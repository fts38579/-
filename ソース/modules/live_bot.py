# -*- coding: utf-8 -*-
"""
カワウソマネージャー きなこ – LiveBot  v9.0
変更:
  v8.8: _gift_last を __init__ に移動（競合リスク解消）
  v8.9: stop_event 連携追加（GUI停止ボタンでループ終了）
  v9.0: [重要修正] client.start() → client.connect() に変更
        start() は非ブロッキング Task を返すだけで WebSocket が動かない
        connect() は接続が切れるまでブロックする正しい API
        gift.name の取得方法を修正（event.gift.name）
        ログ出力を詳細化してデバッグしやすくした
"""

import sys
import os
import time
import asyncio
import threading
import traceback
import csv
from typing import Optional, Callable

from TikTokLive import TikTokLiveClient
from TikTokLive.events import ConnectEvent, DisconnectEvent, GiftEvent, JoinEvent

# ── 定数 ──────────────────────────────────────────────────────────
_OFFLINE_SEC     = 30
_BLOCKED_SEC     = 600
_RETRY_BASE_SEC  = 5
_RETRY_MAX_SEC   = 120
_MAX_RETRIES     = 5
_GIFT_DEDUP_SEC  = 10

# ── プロジェクトルート ────────────────────────────────────────────
if getattr(sys, "frozen", False):
    _PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    _PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def _data_path(rel: str) -> str:
    return os.path.join(_PROJECT_ROOT, rel)

import config

# ── ユーティリティ ────────────────────────────────────────────────
def _safe_str(v) -> str:
    try:    return str(v) if v is not None else ""
    except: return ""

def _extract_user(event):
    """イベントからユーザー名・UID を安全に取得"""
    try:
        u    = event.user
        name = _safe_str(getattr(u, "display_name", "") or getattr(u, "nickname", ""))
        uid  = _safe_str(getattr(u, "unique_id",   "") or getattr(u, "uniqueId",  ""))
        return name or "不明", uid or "不明"
    except:
        return "不明", "不明"

# ── エラー分類 ────────────────────────────────────────────────────
def _is_offline_error(e: Exception) -> bool:
    msgs = ("hosting", "offline", "is not online", "is not live",
            "not currently live", "UserOffline", "LIVE_NOT_FOUND",
            "userofflineerror", "usernotfounderror")
    s = str(e).lower()
    name = type(e).__name__.lower()
    return any(m.lower() in s for m in msgs) or name in msgs

def _is_blocked_error(e: Exception) -> bool:
    names = ("WebcastBlocked200Error", "DeviceBlocked", "DEVICE_BLOCKED")
    s = str(e)
    return any(n in s for n in names) or type(e).__name__ in names

def _is_rate_limit_error(e: Exception) -> bool:
    if _is_blocked_error(e):
        return False
    names = ("RateLimitError", "TooManyRequests", "rate_limit", "RateLimit")
    s = str(e)
    return any(n in s for n in names) or type(e).__name__ in names

# ── カウントダウン付きスリープ ────────────────────────────────────
async def _sleep_cd(seconds: int, label: str,
                    stop_event: Optional[threading.Event] = None):
    end      = time.time() + seconds
    log_step = 60 if seconds >= 600 else 10
    next_log = 0.0
    while True:
        remaining = end - time.time()
        if remaining <= 0:
            break
        if stop_event and stop_event.is_set():
            print(f"[LiveBot] ⏭ {label} – 停止要求により待機をスキップ")
            break
        if time.time() >= next_log:
            m, s = divmod(int(remaining), 60)
            print(f"[LiveBot] ⏳ {label} – 残り {m}分{s:02d}秒")
            next_log = time.time() + log_step
        await asyncio.sleep(min(5, max(0.1, remaining)))

# ── CSV（ギフトタイムライン） ────────────────────────────────────
_CSV_FILE    = _data_path(getattr(config, "CSV_FILE", "data/gift_timeline.csv"))
_CSV_HEADERS = ["timestamp", "type", "user", "unique_id", "detail"]

def _init_csv():
    os.makedirs(os.path.dirname(_CSV_FILE), exist_ok=True)
    if not os.path.exists(_CSV_FILE):
        with open(_CSV_FILE, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(_CSV_HEADERS)

def _append_csv(row_type, user, uid, detail):
    try:
        with open(_CSV_FILE, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow([
                time.strftime("%Y-%m-%d %H:%M:%S"),
                row_type, user, uid, detail
            ])
    except Exception as e:
        print(f"[CSV] 書き込みエラー: {e}")

# ── viewers.csv（入室ログ） ───────────────────────────────────────
_VIEWERS_FILE    = _data_path(getattr(config, "VIEWERS_FILE", "data/viewers.csv"))
_VIEWERS_HEADERS = ["session_date", "session_start", "unique_id", "display_name"]

def _init_viewers_csv():
    os.makedirs(os.path.dirname(_VIEWERS_FILE), exist_ok=True)
    if not os.path.exists(_VIEWERS_FILE):
        with open(_VIEWERS_FILE, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(_VIEWERS_HEADERS)

def _append_viewer(session_date: str, session_start: str, uid: str, name: str):
    try:
        with open(_VIEWERS_FILE, "a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow([session_date, session_start, uid, name])
    except Exception as e:
        print(f"[Viewers] 書き込みエラー: {e}")

# ── リピート率計算 ───────────────────────────────────────────────
def _calc_repeat_rate() -> tuple:
    if not os.path.exists(_VIEWERS_FILE):
        return 0, 0, 0.0
    try:
        uid_sessions: dict = {}
        with open(_VIEWERS_FILE, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                uid  = row.get("unique_id", "").strip()
                date = row.get("session_date", "").strip()
                if uid and uid != "不明" and date:
                    uid_sessions.setdefault(uid, set()).add(date)
        total   = len(uid_sessions)
        repeats = sum(1 for s in uid_sessions.values() if len(s) >= 2)
        rate    = (repeats / total * 100) if total > 0 else 0.0
        return total, repeats, rate
    except Exception as e:
        print(f"[Repeat] 計算エラー: {e}")
        return 0, 0, 0.0

# ── LiveBot ───────────────────────────────────────────────────────
class LiveBot:
    def __init__(self,
                 on_stream_end_callback: Optional[Callable] = None,
                 stop_event: Optional[threading.Event] = None):
        self.username          = config.MY_TIKTOK_USERNAME
        self._on_stream_end_cb = on_stream_end_callback
        self._stop_event       = stop_event  # GUI 停止ボタンと連携
        self.client            = None

        self._stream_started   = False
        self._stream_end_fired = False
        self._should_stop      = False
        self._end_cb_thread: Optional[threading.Thread] = None

        self._start_time         = None
        self._session_date       = ""
        self._session_start_str  = ""
        self._gift_last: dict    = {}

        _init_csv()
        _init_viewers_csv()
        print(f"[LiveBot] 初期化完了 – 監視対象: @{self.username}")

    def _is_stop_requested(self) -> bool:
        if self._stop_event and self._stop_event.is_set():
            return True
        return self._should_stop

    def _fire_end_callback(self):
        """on_stream_end_callback をスレッドで安全に一度だけ発火"""
        if self._on_stream_end_cb:
            if self._end_cb_thread is None or not self._end_cb_thread.is_alive():
                self._end_cb_thread = threading.Thread(
                    target=self._on_stream_end_cb, daemon=False)
                self._end_cb_thread.start()

    # ── イベントハンドラ ──────────────────────────────────────────

    async def _on_connect(self, event: ConnectEvent):
        self._start_time        = time.time()
        self._stream_end_fired  = False
        self._stream_started    = True
        self._session_date      = time.strftime("%Y-%m-%d")
        self._session_start_str = time.strftime("%H:%M:%S")
        print(f"[LiveBot] ✅ 配信開始検知: {self._session_start_str}")
        _append_csv("connect", self.username, "", "配信開始")

    async def _on_disconnect(self, event: DisconnectEvent):
        if self._stream_end_fired:
            return
        self._stream_end_fired = True

        duration = ""
        if self._start_time:
            s = int(time.time() - self._start_time)
            duration = f"{s // 60}分{s % 60}秒"

        print(f"[LiveBot] 📴 配信終了検知 ({duration}): {time.strftime('%H:%M:%S')}")
        _append_csv("disconnect", self.username, "", f"配信終了 {duration}")

        total, repeats, rate = _calc_repeat_rate()
        print("=" * 50)
        print(f"[リピート率] 累計ユニーク視聴者: {total}人")
        print(f"[リピート率] リピーター(2回以上): {repeats}人")
        print(f"[リピート率] リピート率: {rate:.1f}%")
        print("=" * 50)

        self._should_stop = True
        self._fire_end_callback()

    async def _on_gift(self, event: GiftEvent):
        try:
            name, uid = _extract_user(event)
            # v9.0 修正: event.gift.name が正しいアクセス方法
            try:
                gift_name = _safe_str(event.gift.name) if event.gift else "不明"
            except Exception:
                gift_name = _safe_str(getattr(event, "gift_name", "不明"))
            if not gift_name:
                gift_name = "不明"

            # ストリーク中（コンボ継続中）は repeat_end が False = まだ続いている
            # repeat_end が True = ストリーク終了 → その時だけカウント
            if hasattr(event, 'streaking') and event.streaking:
                return  # ストリーク継続中はスキップ（重複防止）

            count = getattr(event, "repeat_count", 1) or 1

            now    = time.time()
            key    = (uid, gift_name)
            last_t = self._gift_last.get(key, 0)
            if now - last_t < _GIFT_DEDUP_SEC:
                return
            self._gift_last[key] = now

            print(f"[Gift] 🎁 {name} が {gift_name} ×{count} を送りました")
            _append_csv("gift", name, uid, f"{gift_name} ×{count}")

        except Exception as e:
            print(f"[Gift] 処理エラー: {e}")
            traceback.print_exc()

    async def _on_join(self, event: JoinEvent):
        try:
            name, uid = _extract_user(event)
            print(f"[Join] 👋 {name} が入室しました")
            _append_csv("join", name, uid, "入室")
            if uid and uid != "不明":
                _append_viewer(
                    self._session_date,
                    self._session_start_str,
                    uid, name
                )
        except Exception as e:
            print(f"[Join] 処理エラー: {e}")

    # ── メインループ ──────────────────────────────────────────────

    async def start(self):
        retry    = 0
        rl_count = 0

        while True:
            # ── 停止チェック ──────────────────────────────────────
            if self._is_stop_requested():
                print("[LiveBot] 🛑 停止要求 – 監視ループを終了します")
                if self._stream_started and not self._stream_end_fired:
                    self._stream_end_fired = True
                    _append_csv("disconnect", self.username, "", "手動停止")
                    self._fire_end_callback()
                break

            print(f"[LiveBot] 🔄 @{self.username} への接続を試みます (試行 {retry + 1})")
            self._stream_started   = False
            self._stream_end_fired = False

            try:
                # ★ v9.0 重要修正: start() → connect() に変更
                # start() は非ブロッキングで Task を返すだけ
                # connect() は WebSocket が切れるまでブロックする
                self.client = TikTokLiveClient(unique_id=f"@{self.username}")
                self.client.add_listener(ConnectEvent,    self._on_connect)
                self.client.add_listener(DisconnectEvent, self._on_disconnect)
                self.client.add_listener(GiftEvent,       self._on_gift)
                self.client.add_listener(JoinEvent,       self._on_join)

                print(f"[LiveBot] 接続中… (@{self.username})")
                await self.client.connect()  # ← ここが修正点
                print(f"[LiveBot] 接続終了 (@{self.username})")
                rl_count = 0

            except Exception as e:
                err  = str(e)
                ename = type(e).__name__

                if _is_rate_limit_error(e):
                    rl_count += 1
                    wait = min(1800 * (2 ** (rl_count - 1)), 7200)
                    if "account_hour" in err:  wait = 3600
                    elif "room_id_day" in err: wait = 7200
                    print(f"[LiveBot] ⏳ レートリミット ({rl_count}回目) – {wait}秒待機")
                    retry = 0
                    await _sleep_cd(wait, "レートリミット待機", self._stop_event)

                elif _is_blocked_error(e):
                    rl_count = 0
                    print(f"[LiveBot] 🚫 ブロック ({ename}: {err[:80]}) – {_BLOCKED_SEC}秒待機")
                    retry = 0
                    await _sleep_cd(_BLOCKED_SEC, "ブロック待機", self._stop_event)

                elif _is_offline_error(e):
                    rl_count = 0
                    print(f"[LiveBot] 📴 @{self.username} はオフラインです – {_OFFLINE_SEC}秒後に再確認")
                    retry = 0
                    await _sleep_cd(_OFFLINE_SEC, "オフライン待機", self._stop_event)

                else:
                    rl_count = 0
                    print(f"[LiveBot] ❌ 予期しないエラー ({ename}): {err[:120]}")
                    traceback.print_exc()

                    if self._stream_started and not self._stream_end_fired:
                        # 配信中に例外 → 配信終了として処理
                        self._stream_end_fired = True
                        duration = ""
                        if self._start_time:
                            s = int(time.time() - self._start_time)
                            duration = f"{s // 60}分{s % 60}秒"
                        _append_csv("disconnect", self.username, "", f"配信終了(例外) {duration}")
                        total, repeats, rate = _calc_repeat_rate()
                        print(f"[リピート率] {total}人中 {repeats}人リピーター ({rate:.1f}%)")
                        self._fire_end_callback()
                        self._should_stop = True
                    else:
                        retry += 1
                        wait = min(_RETRY_BASE_SEC * (2 ** (retry - 1)), _RETRY_MAX_SEC)
                        print(f"[LiveBot] リトライ {retry}/{_MAX_RETRIES} – {wait}秒後")
                        if retry >= _MAX_RETRIES:
                            retry = 0
                            print(f"[LiveBot] 最大リトライ到達 – しばらく待機")
                            await _sleep_cd(60, "リトライ超過待機", self._stop_event)
                        else:
                            await _sleep_cd(wait, "リトライ待機", self._stop_event)

            finally:
                try:
                    if self.client:
                        await self.client.disconnect()
                except Exception:
                    pass
                self.client = None

            # _on_disconnect で _should_stop が立った場合もここで終了
            if self._should_stop:
                print("[LiveBot] ✅ 配信終了 – 監視ループ終了")
                break

        print("[LiveBot] 監視終了")
