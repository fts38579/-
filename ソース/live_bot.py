# -*- coding: utf-8 -*-
"""
カワウソマネージャー きなこ – LiveBot  v8.8
変更: _gift_last を __init__ に移動（競合リスク解消）
"""

import sys
import os
import time
import asyncio
import threading
import traceback
import csv
import queue
from typing import Optional, Callable

from TikTokLive import TikTokLiveClient
from TikTokLive.events import ConnectEvent, DisconnectEvent, GiftEvent, JoinEvent

# ── 定数 ──────────────────────────────────────────────────────────
_MIN_LOOP_SEC    = 10
_STREAM_END_SEC  = 30
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
            "not currently live", "UserOffline", "LIVE_NOT_FOUND")
    s = str(e).lower()
    return any(m.lower() in s for m in msgs) or type(e).__name__ in msgs

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
async def _sleep_cd(seconds: int, label: str):
    end      = time.time() + seconds
    log_step = 60 if seconds >= 600 else 10
    next_log = 0.0
    while True:
        remaining = end - time.time()
        if remaining <= 0:
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
    def __init__(self, on_stream_end_callback: Optional[Callable] = None):
        self.username          = config.MY_TIKTOK_USERNAME
        self._on_stream_end_cb = on_stream_end_callback
        self.client            = None

        self._stream_started   = False
        self._stream_end_fired = False
        self._should_stop      = False
        self._end_cb_thread: Optional[threading.Thread] = None
        self._disconnect_event: Optional[asyncio.Event] = None

        self._start_time         = None
        self._session_date       = ""
        self._session_start_str  = ""

        # ★ _gift_last を __init__ で初期化（競合リスク解消）
        self._gift_last: dict    = {}

        _init_csv()
        _init_viewers_csv()
        print(f"[LiveBot] 初期化完了 – 監視対象: @{self.username}")

    # ── イベント ──────────────────────────────────────────────────

    async def _on_connect(self, event: ConnectEvent):
        self._start_time        = time.time()
        self._stream_end_fired  = False
        self._stream_started    = True
        self._session_date      = time.strftime("%Y-%m-%d")
        self._session_start_str = time.strftime("%H:%M:%S")
        print(f"[LiveBot] ✅ 配信開始: {self._session_start_str}")
        _append_csv("connect", self.username, "", "配信開始")

    async def _on_disconnect(self, event: DisconnectEvent):
        if self._disconnect_event and not self._disconnect_event.is_set():
            self._disconnect_event.set()

        if self._stream_end_fired:
            return
        self._stream_end_fired = True

        duration = ""
        if self._start_time:
            s = int(time.time() - self._start_time)
            duration = f"{s // 60}分{s % 60}秒"

        print(f"[LiveBot] 📴 配信終了 ({duration}): {time.strftime('%H:%M:%S')}")
        _append_csv("disconnect", self.username, "", f"配信終了 {duration}")

        # リピート率を計算して表示
        total, repeats, rate = _calc_repeat_rate()
        print("=" * 50)
        print(f"[リピート率] 累計ユニーク視聴者: {total}人")
        print(f"[リピート率] リピーター(2回以上): {repeats}人")
        print(f"[リピート率] リピート率: {rate:.1f}%")
        print("=" * 50)

        self._should_stop = True

        if self._on_stream_end_cb:
            if self._end_cb_thread is None or not self._end_cb_thread.is_alive():
                self._end_cb_thread = threading.Thread(
                    target=self._on_stream_end_cb, daemon=False)
                self._end_cb_thread.start()

    async def _on_gift(self, event: GiftEvent):
        try:
            name, uid = _extract_user(event)
            gift_name = _safe_str(
                getattr(event, "gift_name", "") or
                getattr(getattr(event, "gift", None), "name", "不明"))
            count = getattr(event, "gift_count", 1) or 1

            now    = time.time()
            key    = (uid, gift_name)
            last_t = self._gift_last.get(key, 0)  # ★ hasattr チェック不要に
            if now - last_t < _GIFT_DEDUP_SEC:
                return
            self._gift_last[key] = now

            print(f"[Gift] 🎁 {name} が {gift_name} ×{count} を送りました")
            _append_csv("gift", name, uid, f"{gift_name} ×{count}")

        except Exception as e:
            print(f"[Gift] 処理エラー: {e}")

    async def _on_join(self, event: JoinEvent):
        try:
            name, uid = _extract_user(event)
            print(f"[Join] 👋 {name} が入室しました")
            _append_csv("join", name, uid, "入室")

            # viewers.csv に記録
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
        retry           = 0
        last_loop_start = 0.0
        rl_count        = 0

        while True:
            elapsed = time.time() - last_loop_start
            if elapsed < _MIN_LOOP_SEC:
                await asyncio.sleep(_MIN_LOOP_SEC - elapsed)
            last_loop_start = time.time()

            if self._should_stop:
                print("[LiveBot] ✅ 配信終了 – 監視終了（再配信は exe 再起動）")
                break

            print(f"[LiveBot] 🔄 @{self.username} への接続を試みます (試行 {retry + 1})")
            self._stream_started   = False
            self._stream_end_fired = False
            self._disconnect_event = asyncio.Event()

            try:
                self.client = TikTokLiveClient(unique_id=f"@{self.username}")
                self.client.add_listener(ConnectEvent,    self._on_connect)
                self.client.add_listener(DisconnectEvent, self._on_disconnect)
                self.client.add_listener(GiftEvent,       self._on_gift)
                self.client.add_listener(JoinEvent,       self._on_join)

                await self.client.start()

                if not self._disconnect_event.is_set():
                    try:
                        await asyncio.wait_for(
                            self._disconnect_event.wait(), timeout=10800)
                    except asyncio.TimeoutError:
                        print("[LiveBot] ⚠ 接続タイムアウト (3時間)")

                rl_count = 0

            except Exception as e:
                err = str(e)
                if self._disconnect_event and not self._disconnect_event.is_set():
                    self._disconnect_event.set()

                if _is_rate_limit_error(e):
                    rl_count += 1
                    wait = min(1800 * (2 ** (rl_count - 1)), 7200)
                    if "account_hour" in err:  wait = 3600
                    elif "room_id_day" in err: wait = 7200
                    print(f"[LiveBot] ⏳ レートリミット ({rl_count}回目)")
                    retry = 0
                    await _sleep_cd(wait, "レートリミット待機")

                elif _is_blocked_error(e):
                    rl_count = 0
                    print(f"[LiveBot] 🚫 ブロック ({err[:80]})")
                    retry = 0
                    await _sleep_cd(_BLOCKED_SEC, "ブロック待機")

                elif _is_offline_error(e):
                    rl_count = 0
                    print(f"[LiveBot] 📴 配信オフライン")
                    retry = 0
                    await _sleep_cd(_OFFLINE_SEC, "オフライン待機")

                else:
                    rl_count = 0
                    if self._stream_started:
                        print(f"[LiveBot] ⚠ 配信中に例外 → 配信終了として処理: {err[:80]}")
                        if not self._stream_end_fired:
                            self._stream_end_fired = True
                            duration = ""
                            if self._start_time:
                                s = int(time.time() - self._start_time)
                                duration = f"{s // 60}分{s % 60}秒"
                            _append_csv("disconnect", self.username, "", f"配信終了(例外) {duration}")
                            total, repeats, rate = _calc_repeat_rate()
                            print("=" * 50)
                            print(f"[リピート率] 累計ユニーク視聴者: {total}人")
                            print(f"[リピート率] リピーター(2回以上): {repeats}人")
                            print(f"[リピート率] リピート率: {rate:.1f}%")
                            print("=" * 50)
                            if self._on_stream_end_cb:
                                if self._end_cb_thread is None or not self._end_cb_thread.is_alive():
                                    self._end_cb_thread = threading.Thread(
                                        target=self._on_stream_end_cb, daemon=False)
                                    self._end_cb_thread.start()
                        self._should_stop = True
                    else:
                        retry += 1
                        wait = min(_RETRY_BASE_SEC * (2 ** (retry - 1)), _RETRY_MAX_SEC)
                        print(f"[LiveBot] ❌ エラー ({retry}/{_MAX_RETRIES}) {wait}s: {e}")
                        traceback.print_exc()
                        if retry >= _MAX_RETRIES:
                            retry = 0
                            await _sleep_cd(_STREAM_END_SEC, "リトライ待機")
                        else:
                            await _sleep_cd(wait, "リトライ待機")

            finally:
                try:
                    if self.client:
                        await self.client.stop()
                except Exception:
                    pass
                self.client = None

            if self._should_stop:
                print("[LiveBot] ✅ 配信終了 – 監視終了（再配信は exe 再起動）")
                break
