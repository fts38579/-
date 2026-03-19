# -*- coding: utf-8 -*-
"""
カワウソマネージャー きなこ – 統合アプリ  v1.0
=============================================
以下の全機能を1つのウィンドウに統合：
  ① 初期セットアップ  （TikTok ID / インサイトURL 設定）
  ② ライブ監視         （メインボット起動 / 停止）
  ③ インサイト手動取得  （即時取得ボタン）
  ④ きなこのレポート    （インサイト / ギフト / リピート率グラフ）
"""

import os
import sys
import re
import time
import threading
import traceback
import io
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ── プロジェクトルート解決 ────────────────────────────────────────
if getattr(sys, "frozen", False):
    _PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    _PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# ソースフォルダも追加（非frozenのみ）
if not getattr(sys, "frozen", False):
    _SRC = os.path.join(_PROJECT_ROOT, "ソース")
    if _SRC not in sys.path:
        sys.path.insert(0, _SRC)

# ── データフォルダ・ファイルパス ─────────────────────────────────
DATA_DIR     = os.path.join(_PROJECT_ROOT, "data")
CSV_FILE     = os.path.join(DATA_DIR, "gift_timeline.csv")
VIEWERS_FILE = os.path.join(DATA_DIR, "viewers.csv")
CONFIG_FILE  = os.path.join(_PROJECT_ROOT, "config.py")

# ── tkcalendar（任意） ───────────────────────────────────────────
try:
    from tkcalendar import DateEntry
    _HAS_CALENDAR = True
except ImportError:
    _HAS_CALENDAR = False

# ── matplotlib ───────────────────────────────────────────────────
try:
    import matplotlib
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import pandas as pd
    _HAS_CHART = True

    def _set_japanese_font():
        candidates = ["Meiryo", "MS Gothic", "Yu Gothic", "IPAGothic",
                      "Noto Sans CJK JP", "TakaoGothic", "IPAPGothic"]
        for name in candidates:
            for f in fm.fontManager.ttflist:
                if name.lower() in f.name.lower():
                    matplotlib.rcParams["font.family"] = f.name
                    return
        for f in fm.fontManager.ttflist:
            if any(c in f.name for c in
                   ["Gothic", "Mincho", "Hiragino", "Noto", "Takao"]):
                matplotlib.rcParams["font.family"] = f.name
                return

    _set_japanese_font()
    matplotlib.rcParams["axes.unicode_minus"] = False

except ImportError:
    _HAS_CHART = False

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  ユーティリティ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Chrome 検索パス
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        r"Google\Chrome\Application\chrome.exe"
    ),
]
ALLOWED_URL_PREFIX = "https://livecenter.tiktok.com/"


def find_chrome() -> str | None:
    for p in CHROME_PATHS:
        if p and os.path.isfile(p):
            return p
    return None


def validate_tiktok_id(value: str) -> str | None:
    if not value:
        return "TikTok ID を入力してください。"
    if len(value) > 24:
        return "TikTok ID は 24 文字以内で入力してください。"
    if not re.fullmatch(r"[a-zA-Z0-9_.]{1,24}", value):
        return ("TikTok ID に使えない文字が含まれています。\n"
                "（使用可能: 英数字・アンダースコア・ピリオドのみ）")
    return None


def validate_url(value: str) -> str | None:
    if not value:
        return "インサイトページ URL を入力してください。"
    if not value.startswith(ALLOWED_URL_PREFIX):
        return (f"URL は以下で始まる TikTok のアドレスのみ使用できます。\n\n"
                f"  {ALLOWED_URL_PREFIX}\n\n"
                "デフォルト値:\n"
                "  https://livecenter.tiktok.com/analytics/live_video?lang=ja-JP")
    return None


def read_config_value(key: str) -> str:
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                m = re.match(rf'^{re.escape(key)}\s*=\s*["\'](.+?)["\']', line)
                if m:
                    return m.group(1)
    except Exception:
        pass
    return ""


def update_config(tiktok_id: str, url: str) -> None:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        content = f.read()

    def replace_var(text: str, key: str, val: str) -> str:
        escaped_key = re.escape(key)
        pattern = rf'^({escaped_key}\s*=\s*)["\'].*?["\']'
        safe_val = repr(val)
        new_text, n = re.subn(
            pattern,
            lambda m: m.group(1) + safe_val,
            text,
            flags=re.MULTILINE
        )
        if n == 0:
            new_text += f"\n{key} = {repr(val)}\n"
        return new_text

    content = replace_var(content, "MY_TIKTOK_USERNAME", tiktok_id)
    content = replace_var(content, "ANALYTICS_URL", url)

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(content)


def find_col(df, *keywords):
    for kw in keywords:
        for col in df.columns:
            if kw in col:
                return col
    return None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  データ読み込み（レポート用）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _insights_csv_path():
    try:
        import config as _cfg
        raw = getattr(_cfg, "CSV_INSIGHTS_FILE", "data/insights.csv")
    except Exception:
        raw = "data/insights.csv"
    return os.path.join(DATA_DIR, os.path.splitext(os.path.basename(raw))[0] + ".csv")


def load_insights():
    path = _insights_csv_path()
    if not os.path.exists(path):
        return None, f"insights.csv が見つかりません。\nパス: {path}"
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
        date_col = None
        if "取得日時" in df.columns:
            date_col = "取得日時"
        else:
            for col in df.columns:
                if any(k in col for k in ["日", "date", "時", "取得"]):
                    date_col = col
                    break
        if date_col:
            df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
        else:
            try:
                df["_date"] = pd.to_datetime(df.iloc[:, 0], errors="coerce")
            except Exception:
                return None, "日付列を認識できませんでした。"
        df = df.dropna(subset=["_date"]).sort_values("_date").reset_index(drop=True)
        return df, None
    except Exception as e:
        return None, str(e)


def load_gifts():
    if not os.path.exists(CSV_FILE):
        return None, "gift_timeline.csv が見つかりません。\ndata フォルダを確認してください。"
    try:
        df = pd.read_csv(CSV_FILE, encoding="utf-8-sig")
        df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
        df = df.dropna(subset=["timestamp"])
        df_gift = df[df["type"] == "gift"].copy()
        if df_gift.empty:
            return df_gift, None

        def parse_detail(s):
            m = re.match(r"^(.+?)\s*×(\d+)$", str(s).strip())
            if m:
                return m.group(1).strip(), int(m.group(2))
            return str(s).strip(), 1

        parsed = df_gift["detail"].apply(parse_detail)
        df_gift = df_gift.copy()
        df_gift["gift_name"] = [p[0] for p in parsed]
        df_gift["count"] = [p[1] for p in parsed]
        df_gift["_date"] = df_gift["timestamp"]
        return df_gift, None
    except Exception as e:
        return None, str(e)


def load_viewers():
    if not os.path.exists(VIEWERS_FILE):
        return None, ("viewers.csv が見つかりません。\n"
                      "配信を1回以上終了すると自動生成されます。")
    try:
        df = pd.read_csv(VIEWERS_FILE, encoding="utf-8-sig")
        df.columns = [c.strip().lower() for c in df.columns]
        if "session_date" not in df.columns:
            df = df.rename(columns={df.columns[0]: "session_date"})
        df["session_date"] = pd.to_datetime(df["session_date"], errors="coerce").dt.date
        df = df.dropna(subset=["session_date"])
        return df, None
    except Exception as e:
        return None, str(e)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  メイン GUI クラス
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class KinakoApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("🦦 カワウソマネージャー きなこ")
        self.geometry("980x780")
        self.configure(bg="#1e1b2e")
        self.resizable(True, True)

        # ライブ監視スレッド管理
        self._bot_thread: threading.Thread | None = None
        self._bot_stop_event = threading.Event()

        # レポートキャッシュ
        self._insight_fig = None
        self._insight_df = None
        self._gift_fig = None
        self._gift_df = None
        self._repeat_fig = None
        self._repeat_df = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ────────────────────────────────────────────────────────────
    #  UI 構築
    # ────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── ヘッダー ──
        hdr = tk.Frame(self, bg="#7c3aed")
        hdr.pack(fill="x")

        tk.Label(
            hdr,
            text="🦦 カワウソマネージャー きなこ",
            font=("Meiryo", 15, "bold"),
            bg="#7c3aed", fg="white",
            padx=16, pady=12
        ).pack(side="left")

        # エクスポートボタン（レポートタブ用）
        self._btn_export_frame = tk.Frame(hdr, bg="#7c3aed")
        self._btn_export_frame.pack(side="right", padx=14, pady=8)

        tk.Button(
            self._btn_export_frame,
            text="📊 Excelエクスポート",
            command=self._on_export_excel,
            font=("Meiryo", 9, "bold"),
            bg="#1d4ed8", fg="white", relief="flat",
            activebackground="#1e3a8a", activeforeground="white",
            padx=10, pady=5
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            self._btn_export_frame,
            text="📥 CSVエクスポート",
            command=self._on_export_csv,
            font=("Meiryo", 9, "bold"),
            bg="#059669", fg="white", relief="flat",
            activebackground="#047857", activeforeground="white",
            padx=10, pady=5
        ).pack(side="left")

        # ── タブ ──
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook", background="#1e1b2e", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=("Meiryo", 11, "bold"),
            padding=[18, 8],
            background="#2d2a45",
            foreground="#c4b5fd"
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#7c3aed")],
            foreground=[("selected", "white")]
        )
        style.configure("TFrame", background="#1e1b2e")

        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill="both", expand=True, padx=10, pady=(8, 0))

        # タブ定義
        self._tab_setup   = ttk.Frame(self._notebook)
        self._tab_live    = ttk.Frame(self._notebook)
        self._tab_insight_get = ttk.Frame(self._notebook)
        self._tab_report  = ttk.Frame(self._notebook)

        self._notebook.add(self._tab_setup,       text="  ⚙️ セットアップ  ")
        self._notebook.add(self._tab_live,         text="  📡 ライブ監視  ")
        self._notebook.add(self._tab_insight_get,  text="  📥 インサイト取得  ")
        self._notebook.add(self._tab_report,       text="  📊 レポート  ")

        self._build_setup_tab(self._tab_setup)
        self._build_live_tab(self._tab_live)
        self._build_insight_get_tab(self._tab_insight_get)
        self._build_report_tab(self._tab_report)

        # ── フッター ──
        tk.Label(
            self,
            text=f"データフォルダ: {DATA_DIR}",
            font=("Meiryo", 8),
            bg="#1e1b2e", fg="#6b7280"
        ).pack(side="bottom", pady=4)

    # ────────────────────────────────────────────────────────────
    #  タブ① セットアップ
    # ────────────────────────────────────────────────────────────

    def _build_setup_tab(self, parent):
        BG = "#1e1b2e"
        parent.configure(style="TFrame")

        wrapper = tk.Frame(parent, bg=BG)
        wrapper.pack(expand=True, fill="both", padx=40, pady=30)

        tk.Label(
            wrapper,
            text="⚙️  初期セットアップ",
            font=("Meiryo", 14, "bold"),
            bg=BG, fg="#c4b5fd"
        ).pack(anchor="w", pady=(0, 4))

        tk.Label(
            wrapper,
            text="2項目を入力して「保存してセットアップ完了」を押してください",
            font=("Meiryo", 10),
            bg=BG, fg="#9ca3af"
        ).pack(anchor="w", pady=(0, 16))

        tk.Label(
            wrapper,
            text=f"📄 config.py の場所: {CONFIG_FILE}",
            font=("Meiryo", 8),
            bg=BG, fg="#6b7280",
            wraplength=860
        ).pack(anchor="w", pady=(0, 14))

        # ── 入力フォーム ──
        def make_row(label_text, default=""):
            tk.Label(
                wrapper, text=label_text,
                bg=BG, fg="#e5e7eb",
                font=("Meiryo", 10, "bold"), anchor="w"
            ).pack(fill="x", pady=(10, 0))
            var = tk.StringVar(value=default)
            entry = tk.Entry(
                wrapper, textvariable=var,
                font=("Meiryo", 11), width=60,
                relief="solid", bd=1,
                bg="#2d2a45", fg="white",
                insertbackground="white"
            )
            entry.pack(fill="x", pady=(3, 0))
            return var

        self._setup_var_id  = make_row(
            "① TikTok ID（@ なし）",
            read_config_value("MY_TIKTOK_USERNAME")
        )
        self._setup_var_url = make_row(
            "② インサイトページ URL",
            read_config_value("ANALYTICS_URL")
        )

        tk.Button(
            wrapper,
            text="✅  保存してセットアップ完了",
            font=("Meiryo", 12, "bold"),
            bg="#7c3aed", fg="white",
            activebackground="#6d28d9", activeforeground="white",
            relief="flat", cursor="hand2",
            padx=16, pady=10,
            command=self._on_setup_save
        ).pack(fill="x", pady=(24, 0))

        tk.Label(
            wrapper,
            text="設定は config.py に保存　／　Chrome は自動検出・永続プロファイルで起動",
            font=("Meiryo", 8),
            bg=BG, fg="#6b7280"
        ).pack(pady=(10, 0))

    def _on_setup_save(self):
        tiktok_id = self._setup_var_id.get().strip().lstrip("@")
        url       = self._setup_var_url.get().strip()

        for err in (validate_tiktok_id(tiktok_id), validate_url(url)):
            if err:
                messagebox.showwarning("入力エラー", err, parent=self)
                return

        if not find_chrome():
            messagebox.showerror(
                "Chrome が見つかりません",
                "Google Chrome がインストールされているか確認してください。\n\n"
                "通常のインストール先:\n"
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                parent=self
            )
            return

        try:
            update_config(tiktok_id, url)
            messagebox.showinfo(
                "セットアップ完了 🎉",
                f"設定を保存しました！\n\n"
                f"  ・TikTok ID : @{tiktok_id}\n\n"
                f"保存先:\n{CONFIG_FILE}\n\n"
                "──────────────────────────────\n"
                "「📡 ライブ監視」タブから配信監視を開始できます！\n\n"
                "※ 初回インサイト取得時に Chrome が起動します。\n"
                "   TikTok にログインすると次回以降は自動ログインです。",
                parent=self
            )
        except PermissionError:
            messagebox.showerror(
                "書き込みエラー",
                f"config.py に書き込めませんでした。\n\n"
                f"対象ファイル:\n{CONFIG_FILE}\n\n"
                "ファイルが他のアプリで開かれていないか確認してください。",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("エラー", f"予期しないエラーが発生しました。\n\n{e}",
                                 parent=self)

    # ────────────────────────────────────────────────────────────
    #  タブ② ライブ監視
    # ────────────────────────────────────────────────────────────

    def _build_live_tab(self, parent):
        BG = "#1e1b2e"

        wrapper = tk.Frame(parent, bg=BG)
        wrapper.pack(expand=True, fill="both", padx=40, pady=30)

        tk.Label(
            wrapper,
            text="📡  ライブ監視ボット",
            font=("Meiryo", 14, "bold"),
            bg=BG, fg="#c4b5fd"
        ).pack(anchor="w", pady=(0, 6))

        tk.Label(
            wrapper,
            text="「監視開始」を押すとバックグラウンドで TikTok ライブを監視します。\n"
                 "配信終了を検知すると、3分後にインサイトを自動取得します。",
            font=("Meiryo", 10),
            bg=BG, fg="#9ca3af",
            justify="left"
        ).pack(anchor="w", pady=(0, 20))

        # ステータス表示
        status_frame = tk.Frame(wrapper, bg="#2d2a45", relief="solid", bd=1)
        status_frame.pack(fill="x", pady=(0, 16))

        tk.Label(
            status_frame,
            text="ステータス:",
            font=("Meiryo", 10, "bold"),
            bg="#2d2a45", fg="#9ca3af",
            padx=12, pady=8
        ).pack(side="left")

        self._live_status_var = tk.StringVar(value="⏹ 停止中")
        self._live_status_label = tk.Label(
            status_frame,
            textvariable=self._live_status_var,
            font=("Meiryo", 11, "bold"),
            bg="#2d2a45", fg="#ef4444",
            padx=4, pady=8
        )
        self._live_status_label.pack(side="left")

        # ボタン
        btn_frame = tk.Frame(wrapper, bg=BG)
        btn_frame.pack(fill="x", pady=(0, 16))

        self._btn_start_live = tk.Button(
            btn_frame,
            text="▶  監視開始",
            font=("Meiryo", 12, "bold"),
            bg="#16a34a", fg="white",
            activebackground="#15803d", activeforeground="white",
            relief="flat", cursor="hand2",
            padx=20, pady=10,
            command=self._on_live_start
        )
        self._btn_start_live.pack(side="left", padx=(0, 12))

        self._btn_stop_live = tk.Button(
            btn_frame,
            text="⏹  監視停止",
            font=("Meiryo", 12, "bold"),
            bg="#dc2626", fg="white",
            activebackground="#b91c1c", activeforeground="white",
            relief="flat", cursor="hand2",
            padx=20, pady=10,
            state="disabled",
            command=self._on_live_stop
        )
        self._btn_stop_live.pack(side="left")

        # ログ表示エリア
        tk.Label(
            wrapper,
            text="ログ",
            font=("Meiryo", 10, "bold"),
            bg=BG, fg="#9ca3af"
        ).pack(anchor="w", pady=(8, 2))

        log_frame = tk.Frame(wrapper, bg=BG)
        log_frame.pack(fill="both", expand=True)

        self._live_log = tk.Text(
            log_frame,
            font=("Consolas", 9),
            bg="#0f0d1a", fg="#d1d5db",
            relief="solid", bd=1,
            wrap="word",
            state="disabled"
        )
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical",
                                  command=self._live_log.yview)
        self._live_log.configure(yscrollcommand=scrollbar.set)
        self._live_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _log(self, msg: str):
        """ライブ監視ログにメッセージを追記（スレッドセーフ）"""
        def _do():
            self._live_log.configure(state="normal")
            self._live_log.insert("end",
                                   f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
            self._live_log.see("end")
            self._live_log.configure(state="disabled")
        self.after(0, _do)

    def _on_live_start(self):
        # config バリデーション
        try:
            import config
            if not config.MY_TIKTOK_USERNAME:
                raise ValueError("MY_TIKTOK_USERNAME が未設定です")
            if not config.ANALYTICS_URL.startswith("https://livecenter.tiktok.com/"):
                raise ValueError("ANALYTICS_URL が不正です")
        except Exception as e:
            messagebox.showerror(
                "セットアップ未完了",
                f"設定に問題があります。\n\n{e}\n\n"
                "「⚙️ セットアップ」タブで設定してください。",
                parent=self
            )
            return

        self._bot_stop_event.clear()
        self._btn_start_live.configure(state="disabled")
        self._btn_stop_live.configure(state="normal")
        self._live_status_var.set("🟢 監視中")
        self._live_status_label.configure(fg="#22c55e")
        self._log("監視ボットを起動しました")
        self._log(f"監視対象: @{config.MY_TIKTOK_USERNAME}")

        def run_bot():
            try:
                import asyncio
                from modules.live_bot import LiveBot

                def on_stream_end():
                    self._log(f"配信終了を検知。3分後にインサイトを自動取得します…")
                    time.sleep(3 * 60)
                    if self._bot_stop_event.is_set():
                        return
                    self._log("インサイト取得中…")
                    try:
                        from modules.insights import collect_insights
                        ok = collect_insights()
                        if ok:
                            self._log("✅ インサイト取得完了！")
                        else:
                            self._log("❌ インサイト取得失敗。debug_page.html を確認してください。")
                    except Exception as ex:
                        self._log(f"❌ インサイト取得エラー: {ex}")

                bot = LiveBot(on_stream_end_callback=on_stream_end)
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    loop.run_until_complete(bot.start())
                finally:
                    loop.close()
            except Exception as ex:
                self._log(f"❌ ボットエラー: {ex}")
                traceback.print_exc()
            finally:
                def restore():
                    self._btn_start_live.configure(state="normal")
                    self._btn_stop_live.configure(state="disabled")
                    self._live_status_var.set("⏹ 停止中")
                    self._live_status_label.configure(fg="#ef4444")
                    self._log("監視ボットが停止しました")
                self.after(0, restore)

        self._bot_thread = threading.Thread(target=run_bot, daemon=True)
        self._bot_thread.start()

    def _on_live_stop(self):
        self._bot_stop_event.set()
        self._btn_stop_live.configure(state="disabled")
        self._log("停止リクエストを送信しました…")
        self._live_status_var.set("⏸ 停止中…")
        self._live_status_label.configure(fg="#f59e0b")

    # ────────────────────────────────────────────────────────────
    #  タブ③ インサイト手動取得
    # ────────────────────────────────────────────────────────────

    def _build_insight_get_tab(self, parent):
        BG = "#1e1b2e"

        wrapper = tk.Frame(parent, bg=BG)
        wrapper.pack(expand=True, fill="both", padx=40, pady=30)

        tk.Label(
            wrapper,
            text="📥  インサイト手動取得",
            font=("Meiryo", 14, "bold"),
            bg=BG, fg="#c4b5fd"
        ).pack(anchor="w", pady=(0, 6))

        tk.Label(
            wrapper,
            text="ボタンを押すと Chrome が自動起動し、TikTok LiveCenter から\n"
                 "最新の配信インサイトデータを取得して data/insights.csv に保存します。",
            font=("Meiryo", 10),
            bg=BG, fg="#9ca3af",
            justify="left"
        ).pack(anchor="w", pady=(0, 24))

        tk.Button(
            wrapper,
            text="📥  今すぐインサイトを取得",
            font=("Meiryo", 13, "bold"),
            bg="#0369a1", fg="white",
            activebackground="#075985", activeforeground="white",
            relief="flat", cursor="hand2",
            padx=20, pady=12,
            command=self._on_insight_get
        ).pack(anchor="w")

        tk.Label(
            wrapper,
            text="\n【手動取得の使い方】\n"
                 "1. 「⚙️ セットアップ」タブで設定を済ませてください\n"
                 "2. 「今すぐインサイトを取得」ボタンを押します\n"
                 "3. Chrome が自動起動します（初回はTikTokログインが必要）\n"
                 "4. 取得完了後、「📊 レポート」タブでグラフを確認できます",
            font=("Meiryo", 10),
            bg=BG, fg="#9ca3af",
            justify="left"
        ).pack(anchor="w", pady=(24, 0))

    def _on_insight_get(self):
        # config バリデーション
        try:
            import config
            config.validate()
        except Exception as e:
            messagebox.showerror(
                "❌ 設定エラー",
                f"config.py の設定に問題があります。\n\n{e}\n\n"
                "「⚙️ セットアップ」タブで設定してください。",
                parent=self
            )
            return

        result = messagebox.askokcancel(
            "インサイト手動取得",
            "TikTok LiveCenter のインサイトを今すぐ取得します。\n\n"
            "Chrome が自動的に起動し、最新の配信データを\n"
            "data/insights.csv に保存します。\n\n"
            "OK で開始、キャンセルで中断します。",
            parent=self
        )
        if not result:
            return

        def run():
            try:
                from modules.insights import collect_insights
                ok = collect_insights()
                def show():
                    if ok:
                        messagebox.showinfo(
                            "✅ 取得完了",
                            "インサイトの取得が完了しました！\n"
                            "保存先: data/insights.csv\n\n"
                            "「📊 レポート」タブでグラフを確認してください。",
                            parent=self
                        )
                    else:
                        messagebox.showwarning(
                            "⚠️ 取得失敗",
                            "インサイトの取得に失敗しました。\n\n"
                            "・TikTok にログインしているか確認してください。\n"
                            "・data/debug_page.html で詳細を確認できます。",
                            parent=self
                        )
                self.after(0, show)
            except Exception as ex:
                def show_err():
                    messagebox.showerror(
                        "❌ エラー",
                        f"インサイト取得中にエラーが発生しました。\n\n{ex}\n\n"
                        "・Chrome が既に起動していないか確認してください。\n"
                        "・TikTok にログインしているか確認してください。",
                        parent=self
                    )
                self.after(0, show_err)
                traceback.print_exc()

        threading.Thread(target=run, daemon=True).start()

    # ────────────────────────────────────────────────────────────
    #  タブ④ レポート（インサイト / ギフト / リピート率）
    # ────────────────────────────────────────────────────────────

    def _build_report_tab(self, parent):
        if not _HAS_CHART:
            tk.Label(
                parent,
                text="⚠️  matplotlib / pandas がインストールされていません。\n"
                     "  py -m pip install matplotlib pandas",
                font=("Meiryo", 11),
                bg="#1e1b2e", fg="#f59e0b"
            ).pack(expand=True)
            return

        style = ttk.Style()
        style.configure("Report.TNotebook", background="#1e1b2e", borderwidth=0)
        style.configure(
            "Report.TNotebook.Tab",
            font=("Meiryo", 10, "bold"),
            padding=[14, 6],
            background="#2d2a45",
            foreground="#c4b5fd"
        )
        style.map(
            "Report.TNotebook.Tab",
            background=[("selected", "#5b21b6")],
            foreground=[("selected", "white")]
        )

        sub_nb = ttk.Notebook(parent, style="Report.TNotebook")
        sub_nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        self._report_sub_nb = sub_nb

        self._rtab_insight = ttk.Frame(sub_nb)
        self._rtab_gift    = ttk.Frame(sub_nb)
        self._rtab_repeat  = ttk.Frame(sub_nb)

        sub_nb.add(self._rtab_insight, text="  📊 インサイト  ")
        sub_nb.add(self._rtab_gift,    text="  🎁 ギフト  ")
        sub_nb.add(self._rtab_repeat,  text="  👥 リピート率  ")

        self._build_insight_report_tab(self._rtab_insight)
        self._build_gift_report_tab(self._rtab_gift)
        self._build_repeat_report_tab(self._rtab_repeat)

        # 遅延描画
        self.after(300, self._on_show_insights)
        self.after(400, self._on_show_gift)
        self.after(500, self._on_show_repeat)

    # ── 共通ウィジェット ──

    def _make_date_entry(self, parent, var: tk.StringVar):
        if _HAS_CALENDAR:
            today = datetime.today()
            try:
                init_date = datetime.strptime(var.get(), "%Y-%m-%d")
            except Exception:
                init_date = today
            return DateEntry(
                parent,
                textvariable=var,
                font=("Meiryo", 10),
                width=14,
                date_pattern="yyyy-mm-dd",
                year=init_date.year,
                month=init_date.month,
                day=init_date.day,
                background="#7c3aed",
                foreground="white",
                borderwidth=1,
            )
        return tk.Entry(
            parent, textvariable=var,
            font=("Meiryo", 10), width=16,
            relief="solid", bd=1,
            bg="#2d2a45", fg="white",
            insertbackground="white"
        )

    def _embed_chart(self, frame, fig):
        for w in frame.winfo_children():
            w.destroy()
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        return canvas

    def _make_ctrl_row(self, parent, on_show_cmd):
        """日付範囲コントロール行を作成、(var_start, var_end) を返す"""
        today     = datetime.today()
        one_month = today - timedelta(days=30)

        ctrl = tk.Frame(parent, bg="#1e1b2e")
        ctrl.pack(fill="x", padx=16, pady=(10, 4))

        tk.Label(ctrl, text="開始日", bg="#1e1b2e", fg="#d1d5db",
                 font=("Meiryo", 10)).grid(row=0, column=0, padx=(0, 4))
        var_start = tk.StringVar(value=one_month.strftime("%Y-%m-%d"))
        self._make_date_entry(ctrl, var_start).grid(row=0, column=1, padx=(0, 16))

        tk.Label(ctrl, text="終了日", bg="#1e1b2e", fg="#d1d5db",
                 font=("Meiryo", 10)).grid(row=0, column=2, padx=(0, 4))
        var_end = tk.StringVar(value=today.strftime("%Y-%m-%d"))
        self._make_date_entry(ctrl, var_end).grid(row=0, column=3, padx=(0, 16))

        tk.Button(
            ctrl, text="グラフを表示",
            command=on_show_cmd,
            font=("Meiryo", 10, "bold"),
            bg="#7c3aed", fg="white", relief="flat",
            activebackground="#5b21b6", activeforeground="white",
            padx=14, pady=4
        ).grid(row=0, column=4)

        return var_start, var_end

    # ── インサイトレポートタブ ──

    def _build_insight_report_tab(self, parent):
        self._var_ins_start, self._var_ins_end = \
            self._make_ctrl_row(parent, self._on_show_insights)
        self._frame_ins_graph = tk.Frame(parent, bg="#1e1b2e")
        self._frame_ins_graph.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _on_show_insights(self):
        if not _HAS_CHART:
            return
        df, err = load_insights()
        if err:
            return  # データなしは黙って無視
        try:
            start = pd.to_datetime(self._var_ins_start.get())
            end   = (pd.to_datetime(self._var_ins_end.get())
                     + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))
        except Exception:
            return
        df = df[(df["_date"] >= start) & (df["_date"] <= end)]
        if df.empty:
            return

        col_peak    = ("最高同時視聴者数" if "最高同時視聴者数" in df.columns
                       else find_col(df, "最高同時", "peak", "同接"))
        col_diamond = ("ダイヤ合計" if "ダイヤ合計" in df.columns
                       else find_col(df, "diamond"))
        col_gifter  = ("ギフト贈呈者数" if "ギフト贈呈者数" in df.columns
                       else find_col(df, "ギフト贈呈", "gifter"))
        col_watch   = ("平均視聴時間" if "平均視聴時間" in df.columns
                       else find_col(df, "平均視聴", "watch", "view"))

        for col in [col_peak, col_diamond, col_gifter, col_watch]:
            if col:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        diamond_total = int(df[col_diamond].sum(skipna=True)) if col_diamond else 0
        title = (f"インサイト（{self._var_ins_start.get()} "
                 f"～ {self._var_ins_end.get()}）"
                 f"  ◆ 期間合計ダイヤ: {diamond_total:,}")

        plt.close("all")
        fig, axes = plt.subplots(2, 2, figsize=(10, 6),
                                 facecolor="#1e1b2e")
        fig.suptitle(title, fontsize=10, y=0.98, color="#e5e7eb")

        plot_configs = [
            (axes[0][0], col_peak,    "#4f86c6", "最高同時視聴者数（人）"),
            (axes[0][1], col_diamond, "#f5a623", "ダイヤ数"),
            (axes[1][0], col_gifter,  "#7ed321", "ギフト贈呈者数（人）"),
            (axes[1][1], col_watch,   "#e87c7c", "平均視聴時間"),
        ]
        for ax, col, color, ylabel in plot_configs:
            ax.set_facecolor("#2d2a45")
            ax.tick_params(colors="#9ca3af")
            for spine in ax.spines.values():
                spine.set_color("#4b5563")
            if col and col in df.columns:
                mask        = df[col].notna()
                vals        = df.loc[mask, col]
                date_labels = (df.loc[mask, "_date"].dt.strftime("%m/%d")
                               if "_date" in df.columns
                               else [str(i) for i in range(len(vals))])
                if not vals.empty:
                    ax.bar(range(len(vals)), vals, color=color, alpha=0.85)
                    ax.set_xticks(range(len(vals)))
                    ax.set_xticklabels(date_labels, rotation=45,
                                       fontsize=8, color="#9ca3af")
                    ax.set_ylabel(ylabel, fontsize=9, color="#9ca3af")
                    mean_val = vals.mean()
                    ax.axhline(mean_val, color="red", linestyle="--",
                               linewidth=1.2, label=f"平均: {mean_val:.1f}")
                    ax.legend(fontsize=8, facecolor="#2d2a45",
                              labelcolor="#e5e7eb")
                else:
                    ax.text(0.5, 0.5, "データなし", ha="center",
                            va="center", transform=ax.transAxes,
                            color="#9ca3af", fontsize=11)
            else:
                ax.text(0.5, 0.5, "データなし", ha="center",
                        va="center", transform=ax.transAxes,
                        color="#9ca3af", fontsize=11)
            ax.set_title(ylabel, fontsize=10, color="#c4b5fd")

        plt.tight_layout(rect=[0, 0, 1, 0.94])
        self._insight_fig = fig
        self._insight_df  = df
        self._embed_chart(self._frame_ins_graph, fig)

    # ── ギフトレポートタブ ──

    def _build_gift_report_tab(self, parent):
        self._var_gift_start, self._var_gift_end = \
            self._make_ctrl_row(parent, self._on_show_gift)
        self._frame_gift_graph = tk.Frame(parent, bg="#1e1b2e")
        self._frame_gift_graph.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _on_show_gift(self):
        if not _HAS_CHART:
            return
        df, err = load_gifts()
        if err or df is None or df.empty:
            return
        try:
            start = pd.to_datetime(self._var_gift_start.get()).date()
            end   = (pd.to_datetime(self._var_gift_end.get())
                     + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)).date()
        except Exception:
            return
        df_range = df[(df["_date"].dt.date >= start) &
                      (df["_date"].dt.date <= end)].copy()
        if df_range.empty:
            return

        unique_gifters = (df_range["user"].nunique()
                          if "user" in df_range.columns else 0)
        gift_count = len(df_range)
        period_str = f"{self._var_gift_start.get()} ～ {self._var_gift_end.get()}"

        plt.close("all")
        fig, axes = plt.subplots(1, 3, figsize=(13, 5), facecolor="#1e1b2e")
        fig.suptitle(
            f"ギフトタイムライン（{period_str}）"
            f"  ◆ ギフター: {unique_gifters}人  |  ギフト回数: {gift_count}回",
            fontsize=10, y=0.98, color="#e5e7eb"
        )

        axes_configs = [
            (axes[0], "時間帯別ギフト回数"),
            (axes[1], "トップギフター Top10"),
            (axes[2], "ギフト種別 Top10"),
        ]
        for ax, title in axes_configs:
            ax.set_facecolor("#2d2a45")
            ax.tick_params(colors="#9ca3af")
            for spine in ax.spines.values():
                spine.set_color("#4b5563")
            ax.set_title(title, fontsize=10, color="#c4b5fd")

        ax1, ax2, ax3 = axes
        if "_date" in df_range.columns:
            df_range["hour"] = df_range["_date"].dt.hour
            hourly = df_range.groupby("hour").size()
            ax1.bar(hourly.index, hourly.values, color="#f5a623", alpha=0.85)
            ax1.set_xlabel("時刻（時）", fontsize=9, color="#9ca3af")
            ax1.set_ylabel("ギフト回数", fontsize=9, color="#9ca3af")

        if "user" in df_range.columns:
            top_gifters = df_range.groupby("user").size().nlargest(10)
            ax2.barh(top_gifters.index[::-1], top_gifters.values[::-1],
                     color="#7ed321", alpha=0.85)
            ax2.set_xlabel("ギフト回数", fontsize=9, color="#9ca3af")
            ax2.tick_params(axis="y", labelcolor="#d1d5db", labelsize=8)

        if "gift_name" in df_range.columns:
            top_gifts = df_range.groupby("gift_name")["count"].sum().nlargest(10)
            ax3.barh(top_gifts.index[::-1], top_gifts.values[::-1],
                     color="#4f86c6", alpha=0.85)
            ax3.set_xlabel("合計個数", fontsize=9, color="#9ca3af")
            ax3.tick_params(axis="y", labelcolor="#d1d5db", labelsize=8)

        plt.tight_layout(rect=[0, 0, 1, 0.94])
        self._gift_fig = fig
        self._gift_df  = df_range
        self._embed_chart(self._frame_gift_graph, fig)

    # ── リピート率レポートタブ ──

    def _build_repeat_report_tab(self, parent):
        self._var_rep_start, self._var_rep_end = \
            self._make_ctrl_row(parent, self._on_show_repeat)
        self._frame_rep_graph = tk.Frame(parent, bg="#1e1b2e")
        self._frame_rep_graph.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _on_show_repeat(self):
        if not _HAS_CHART:
            return
        df, err = load_viewers()
        if err or df is None or df.empty:
            return
        try:
            start = pd.to_datetime(self._var_rep_start.get()).date()
            end   = (pd.to_datetime(self._var_rep_end.get())
                     + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)).date()
            df = df[(df["session_date"] >= start) & (df["session_date"] <= end)]
        except Exception:
            return
        if df.empty:
            return

        uid_col  = ("uid" if "uid" in df.columns
                    else df.columns[2] if len(df.columns) > 2 else None)
        name_col = "display_name" if "display_name" in df.columns else None
        if uid_col is None:
            return

        session_counts  = df.groupby(uid_col)["session_date"].nunique()
        total_unique    = len(session_counts)
        repeaters       = int((session_counts >= 2).sum())
        first_only      = total_unique - repeaters
        repeat_rate     = (repeaters / total_unique * 100) if total_unique > 0 else 0.0
        session_viewers = df.groupby("session_date")[uid_col].nunique().sort_index()
        top_repeaters   = session_counts[session_counts >= 2].nlargest(10)

        if name_col:
            name_map   = df.drop_duplicates(uid_col).set_index(uid_col)[name_col]
            top_labels = [name_map.get(u, str(u)) for u in top_repeaters.index]
        else:
            top_labels = [str(u) for u in top_repeaters.index]

        plt.close("all")
        fig, axes = plt.subplots(1, 3, figsize=(13, 5), facecolor="#1e1b2e")
        fig.suptitle(
            f"リピート率レポート  |  ユニーク視聴者: {total_unique}人  "
            f"リピーター: {repeaters}人  リピート率: {repeat_rate:.1f}%",
            fontsize=10, y=0.98, color="#e5e7eb"
        )

        for ax in axes:
            ax.set_facecolor("#2d2a45")
            ax.tick_params(colors="#9ca3af")
            for spine in ax.spines.values():
                spine.set_color("#4b5563")

        ax0, ax1, ax2 = axes

        ax0.set_title("リピーター比率", fontsize=10, color="#c4b5fd")
        if total_unique > 0:
            ax0.pie(
                [repeaters, first_only],
                labels=[f"リピーター\n{repeaters}人", f"初回のみ\n{first_only}人"],
                colors=["#7c3aed", "#c4b5fd"],
                autopct="%1.1f%%", startangle=90,
                textprops={"fontsize": 10, "color": "#e5e7eb"}
            )

        ax1.set_title("セッション別ユニーク視聴者", fontsize=10, color="#c4b5fd")
        if not session_viewers.empty:
            dates = [str(d) for d in session_viewers.index]
            ax1.bar(range(len(dates)), session_viewers.values,
                    color="#4f86c6", alpha=0.85)
            ax1.set_xticks(range(len(dates)))
            ax1.set_xticklabels(dates, rotation=45,
                                fontsize=8, color="#9ca3af")
            ax1.set_ylabel("ユニーク視聴者数（人）",
                           fontsize=9, color="#9ca3af")
            mean_v = session_viewers.mean()
            ax1.axhline(mean_v, color="red", linestyle="--", linewidth=1.2,
                        label=f"平均: {mean_v:.1f}")
            ax1.legend(fontsize=8, facecolor="#2d2a45", labelcolor="#e5e7eb")

        ax2.set_title("リピーター Top10（参加セッション数）",
                      fontsize=10, color="#c4b5fd")
        if len(top_repeaters) > 0:
            ax2.barh(top_labels[::-1], top_repeaters.values[::-1],
                     color="#7ed321", alpha=0.85)
            ax2.set_xlabel("参加セッション数", fontsize=9, color="#9ca3af")
            ax2.tick_params(axis="y", labelcolor="#d1d5db", labelsize=8)

        plt.tight_layout(rect=[0, 0, 1, 0.94])
        self._repeat_fig = fig
        self._repeat_df  = df
        self._embed_chart(self._frame_rep_graph, fig)

    # ────────────────────────────────────────────────────────────
    #  エクスポート
    # ────────────────────────────────────────────────────────────

    def _get_current_report(self):
        """現在表示中のレポートサブタブの (fig, df, title) を返す"""
        try:
            idx = self._report_sub_nb.index(self._report_sub_nb.select())
        except Exception:
            return None, None, ""
        if idx == 0:
            title = (f"インサイト（{self._var_ins_start.get()} "
                     f"～ {self._var_ins_end.get()}）")
            return self._insight_fig, self._insight_df, title
        elif idx == 1:
            title = (f"ギフトタイムライン（{self._var_gift_start.get()} "
                     f"～ {self._var_gift_end.get()}）")
            return self._gift_fig, self._gift_df, title
        else:
            return self._repeat_fig, self._repeat_df, "リピート率レポート"

    def _on_export_excel(self):
        fig, df, title = self._get_current_report()
        if fig is None or df is None:
            messagebox.showwarning(
                "未表示",
                "先にレポートタブを開いてグラフを表示してからエクスポートしてください。",
                parent=self
            )
            return
        try:
            import openpyxl
            from openpyxl.drawing.image import Image as XLImage
        except ImportError:
            messagebox.showerror(
                "ライブラリエラー",
                "openpyxl がインストールされていません。\n  py -m pip install openpyxl",
                parent=self
            )
            return

        safe_title = (title.replace("（", "_").replace("）", "")
                           .replace(" ", "").replace("～", "-"))
        path = filedialog.asksaveasfilename(
            title="Excelファイルを保存",
            initialfile=f"{safe_title}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
            parent=self
        )
        if not path:
            return
        try:
            wb      = openpyxl.Workbook()
            ws_data = wb.active
            ws_data.title = "データ"
            export_df = df.copy()
            for col in export_df.columns:
                if pd.api.types.is_datetime64_any_dtype(export_df[col]):
                    export_df[col] = export_df[col].astype(str)
            ws_data.append(list(export_df.columns))
            for row in export_df.itertuples(index=False):
                ws_data.append(list(row))
            for col_cells in ws_data.columns:
                max_len = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col_cells
                )
                ws_data.column_dimensions[
                    col_cells[0].column_letter].width = min(max_len + 4, 40)

            ws_chart  = wb.create_sheet(title="グラフ")
            img_buf   = io.BytesIO()
            fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
            img_buf.seek(0)
            ws_chart.add_image(XLImage(img_buf), "A1")
            wb.save(path)
            messagebox.showinfo("保存完了",
                                f"Excelファイルを保存しました。\n{path}",
                                parent=self)
        except Exception as e:
            messagebox.showerror("保存エラー",
                                 f"Excel保存に失敗しました。\n{e}",
                                 parent=self)

    def _on_export_csv(self):
        fig, df, title = self._get_current_report()
        if df is None:
            messagebox.showwarning(
                "未表示",
                "先にレポートタブを開いてグラフを表示してからエクスポートしてください。",
                parent=self
            )
            return
        safe_title = (title.replace("（", "_").replace("）", "")
                           .replace(" ", "").replace("～", "-"))
        path = filedialog.asksaveasfilename(
            title="CSVを保存",
            initialfile=f"{safe_title}.csv",
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
            parent=self
        )
        if not path:
            return
        try:
            out = df.copy()
            for col in out.columns:
                if pd.api.types.is_datetime64_any_dtype(out[col]):
                    out[col] = out[col].astype(str)
            out.to_csv(path, index=False, encoding="utf-8-sig")
            messagebox.showinfo("保存完了",
                                f"CSVを保存しました。\n{path}",
                                parent=self)
        except Exception as e:
            messagebox.showerror("保存エラー",
                                 f"保存に失敗しました。\n{e}",
                                 parent=self)

    # ────────────────────────────────────────────────────────────
    #  終了処理
    # ────────────────────────────────────────────────────────────

    def _on_close(self):
        if (self._bot_thread and self._bot_thread.is_alive()):
            if not messagebox.askyesno(
                "終了確認",
                "ライブ監視ボットが動作中です。\n終了してよいですか？",
                parent=self
            ):
                return
        self._bot_stop_event.set()
        plt.close("all")
        self.destroy()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  エントリポイント
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == "__main__":
    app = KinakoApp()
    app.mainloop()
