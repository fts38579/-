# -*- coding: utf-8 -*-
"""
カワウソマネージャー きなこ – メインエントリポイント  v8.5
"""

import sys
import os
import time
import threading
import traceback
import asyncio
import msvcrt

# ── プロジェクトルート解決 ─────────────────────────────────────────
if getattr(sys, "frozen", False):
    _PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    _PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# ── 多重起動防止 ───────────────────────────────────────────────────
_LOCK_FILE = os.path.join(_PROJECT_ROOT, ".kawausolock")
_lock_fh   = None

def _acquire_lock() -> bool:
    global _lock_fh
    try:
        _lock_fh = open(_LOCK_FILE, "w")
        msvcrt.locking(_lock_fh.fileno(), msvcrt.LK_NBLCK, 1)
        return True
    except OSError:
        return False

def _release_lock():
    global _lock_fh
    if _lock_fh:
        try:
            msvcrt.locking(_lock_fh.fileno(), msvcrt.LK_UNLCK, 1)
            _lock_fh.close()
        except Exception:
            pass
        try:
            os.remove(_LOCK_FILE)
        except Exception:
            pass
        _lock_fh = None

if not _acquire_lock():
    print("⚠ 既に起動しています。多重起動はできません。")
    input("Enter を押して終了…")
    sys.exit(1)

# ── config 読み込み ────────────────────────────────────────────────
try:
    import config
except ImportError as e:
    print(f"❌ config.py が見つかりません: {e}")
    _release_lock()
    input("Enter を押して終了…")
    sys.exit(1)
except Exception as e:
    print(f"❌ config 読み込みエラー: {e}")
    traceback.print_exc()
    _release_lock()
    input("Enter を押して終了…")
    sys.exit(1)

from modules.live_bot import LiveBot

# ── インサイト自動取得 ─────────────────────────────────────────────
_INSIGHT_DELAY_SEC = 3 * 60   # ★ 3分に変更
_insight_lock      = threading.Lock()
_insight_running   = False

def auto_collect_insights():
    global _insight_running
    with _insight_lock:
        if _insight_running:
            print("[インサイト] 既に実行中のためスキップ")
            return
        _insight_running = True
    try:
        print(f"[インサイト] 配信終了を検知 – {_INSIGHT_DELAY_SEC // 60}分後に自動取得します")
        time.sleep(_INSIGHT_DELAY_SEC)
        print("[インサイト] 取得開始...")
        try:
            from modules.insights import collect_insights
            ok = collect_insights()
            if ok:
                print("✅ [インサイト] 取得完了！data/insights.xlsx に保存しました。")
            else:
                print("❌ [インサイト] 取得失敗。data/debug_page.html を確認してください。")
        except Exception as e:
            print(f"❌ [インサイト] 取得エラー: {e}")
            traceback.print_exc()
    finally:
        with _insight_lock:
            _insight_running = False

# ── メイン ────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("  カワウソマネージャー きなこ 起動")
    print(f"  監視対象: @{config.MY_TIKTOK_USERNAME}")
    print("=" * 60)

    bot = LiveBot(on_stream_end_callback=auto_collect_insights)
    try:
        asyncio.run(bot.start())
    except KeyboardInterrupt:
        print("\n[main] Ctrl+C で終了")
    except Exception as e:
        print(f"[main] 予期しないエラー: {e}")
        traceback.print_exc()
    finally:
        t = getattr(bot, "_end_cb_thread", None)
        if t and t.is_alive():
            print("[main] ⏳ インサイト取得完了まで待機中… (最大3分)")
            t.join(timeout=180)
            if t.is_alive():
                print("[main] ⚠ タイムアウト – インサイト取得未完了")

        print("[main] 終了処理完了")
        _release_lock()
        input("\n[main] Enter を押して閉じる…")
