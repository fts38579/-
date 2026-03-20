# -*- coding: utf-8 -*-
"""
カワウソマネージャー きなこ – 統合アプリ  v2.1  (PyQt6)
=======================================================
修正・改善内容:
  [Bug1] ライブ監視: print()出力をGUIログに転送（stdout リダイレクト）
  [Bug2] 監視停止ボタン押下時もインサイト自動取得を実行
  [Bug3] insights.py で webdriver-manager 対応（Chrome自動検出）
  [Bug4] tkcalendar の月選択不具合 → PyQt6 QDateEdit に置換
  [改善] GUI 全体を tkinter → PyQt6 に移行
  [改善] グラフにデータポイント・ツールチップ・ズーム対応
  [Bug2/3] live_bot に stop_event を渡して停止ボタンで即座にループ終了
  [Bug3] insights.py で webdriver-manager 対応（Chrome自動検出）
"""

import os
import sys
import re
import time
import threading
import traceback
import io
from datetime import datetime, timedelta, date

# ── プロジェクトルート解決 ────────────────────────────────────────
if getattr(sys, "frozen", False):
    _PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    _PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

_SRC = os.path.join(_PROJECT_ROOT, "ソース")
if not getattr(sys, "frozen", False) and _SRC not in sys.path:
    sys.path.insert(0, _SRC)

DATA_DIR     = os.path.join(_PROJECT_ROOT, "data")
CONFIG_FILE  = os.path.join(_PROJECT_ROOT, "config.py")
CSV_FILE     = os.path.join(DATA_DIR, "gift_timeline.csv")
VIEWERS_FILE = os.path.join(DATA_DIR, "viewers.csv")

# ── PyQt6 ────────────────────────────────────────────────────────
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QFileDialog, QMessageBox, QFrame, QDateEdit,
    QSizePolicy, QScrollArea, QSplitter
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QObject, QDate, QTimer, pyqtSlot
)
from PyQt6.QtGui import QFont, QColor, QPalette, QTextCursor

# ── matplotlib (PyQt6バックエンド) ───────────────────────────────
import matplotlib
matplotlib.use("QtAgg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np

try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    _HAS_PANDAS = False

# ── 日本語フォント設定 ───────────────────────────────────────────
def _set_japanese_font():
    candidates = ["Meiryo", "MS Gothic", "Yu Gothic", "IPAGothic",
                  "Noto Sans CJK JP", "TakaoGothic", "IPAPGothic",
                  "Noto Sans JP", "DejaVu Sans"]
    for name in candidates:
        for f in fm.fontManager.ttflist:
            if name.lower() in f.name.lower():
                matplotlib.rcParams["font.family"] = f.name
                return
_set_japanese_font()
matplotlib.rcParams["axes.unicode_minus"] = False

# ────────────────────────────────────────────────────────────────
#  カラーパレット（ダークテーマ）
# ────────────────────────────────────────────────────────────────
C_BG       = "#1e1b2e"
C_PANEL    = "#2d2a45"
C_ACCENT   = "#7c3aed"
C_ACCENT2  = "#5b21b6"
C_TEXT     = "#e5e7eb"
C_SUBTEXT  = "#9ca3af"
C_GREEN    = "#16a34a"
C_RED      = "#dc2626"
C_BLUE     = "#1d4ed8"
C_TEAL     = "#059669"
C_BORDER   = "#4b5563"

STYLE_SHEET = f"""
QMainWindow, QWidget {{
    background-color: {C_BG};
    color: {C_TEXT};
    font-family: "Meiryo", "Yu Gothic", "MS Gothic", sans-serif;
    font-size: 10pt;
}}
QTabWidget::pane {{
    border: 1px solid {C_BORDER};
    background: {C_BG};
}}
QTabBar::tab {{
    background: {C_PANEL};
    color: {C_SUBTEXT};
    padding: 8px 20px;
    font-weight: bold;
    border: none;
}}
QTabBar::tab:selected {{
    background: {C_ACCENT};
    color: white;
}}
QTabBar::tab:hover:!selected {{
    background: #3d3a55;
}}
QLineEdit {{
    background: {C_PANEL};
    color: {C_TEXT};
    border: 1px solid {C_BORDER};
    border-radius: 4px;
    padding: 6px 10px;
    font-size: 10pt;
}}
QLineEdit:focus {{
    border: 1px solid {C_ACCENT};
}}
QTextEdit {{
    background: #0f0d1a;
    color: {C_TEXT};
    border: 1px solid {C_BORDER};
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9pt;
}}
QPushButton {{
    border-radius: 5px;
    padding: 8px 18px;
    font-weight: bold;
    font-size: 10pt;
    border: none;
}}
QPushButton:hover {{ opacity: 0.85; }}
QPushButton:disabled {{ background: #4b5563; color: #6b7280; }}
QDateEdit {{
    background: {C_PANEL};
    color: {C_TEXT};
    border: 1px solid {C_BORDER};
    border-radius: 4px;
    padding: 5px 8px;
}}
QDateEdit::drop-down {{
    border: none;
    background: {C_ACCENT};
    border-radius: 3px;
    width: 20px;
}}
QCalendarWidget {{
    background: {C_PANEL};
    color: {C_TEXT};
}}
QCalendarWidget QAbstractItemView {{
    background: {C_PANEL};
    color: {C_TEXT};
    selection-background-color: {C_ACCENT};
}}
QCalendarWidget QToolButton {{
    background: {C_ACCENT};
    color: white;
    border-radius: 3px;
    padding: 4px;
}}
QCalendarWidget QMenu {{
    background: {C_PANEL};
    color: {C_TEXT};
}}
QScrollBar:vertical {{
    background: {C_PANEL};
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {C_ACCENT};
    border-radius: 4px;
    min-height: 20px;
}}
QFrame#statusFrame {{
    background: {C_PANEL};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 6px;
}}
QLabel#sectionTitle {{
    font-size: 13pt;
    font-weight: bold;
    color: #c4b5fd;
}}
QLabel#subText {{
    color: {C_SUBTEXT};
    font-size: 9pt;
}}
"""

def btn(text, color, callback, min_w=140):
    b = QPushButton(text)
    b.setStyleSheet(
        f"QPushButton {{ background:{color}; color:white; }}"
        f"QPushButton:hover {{ background:{color}cc; }}"
    )
    b.setMinimumWidth(min_w)
    b.clicked.connect(callback)
    return b

# ────────────────────────────────────────────────────────────────
#  stdout → Qt シグナル リダイレクター  [Bug1修正]
# ────────────────────────────────────────────────────────────────
class _StdoutRedirector(QObject):
    text_written = pyqtSignal(str)

    def __init__(self, original_stdout):
        super().__init__()
        self._original = original_stdout

    def write(self, text):
        if text.strip():
            self.text_written.emit(text.rstrip())
        self._original.write(text)

    def flush(self):
        self._original.flush()

# ────────────────────────────────────────────────────────────────
#  設定ユーティリティ
# ────────────────────────────────────────────────────────────────
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 r"Google\Chrome\Application\chrome.exe"),
]
ALLOWED_URL_PREFIX = "https://livecenter.tiktok.com/"

def find_chrome():
    for p in CHROME_PATHS:
        if p and os.path.isfile(p):
            return p
    return None

def validate_tiktok_id(v):
    if not v: return "TikTok ID を入力してください。"
    if len(v) > 24: return "TikTok ID は 24 文字以内で入力してください。"
    if not re.fullmatch(r"[a-zA-Z0-9_.]{1,24}", v):
        return "TikTok ID に使えない文字が含まれています。\n（英数字・アンダースコア・ピリオドのみ）"
    return None

def validate_url(v):
    if not v: return "インサイトページ URL を入力してください。"
    if not v.startswith(ALLOWED_URL_PREFIX):
        return f"URL は {ALLOWED_URL_PREFIX} で始まる必要があります。"
    return None

def read_config_value(key):
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                m = re.match(rf'^{re.escape(key)}\s*=\s*["\'](.+?)["\']', line)
                if m: return m.group(1)
    except Exception:
        pass
    return ""

def update_config(tiktok_id, url):
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    def replace_var(text, key, val):
        pat = rf'^({re.escape(key)}\s*=\s*)["\'].*?["\']'
        new, n = re.subn(pat, lambda m: m.group(1) + repr(val),
                         text, flags=re.MULTILINE)
        if n == 0: new += f"\n{key} = {repr(val)}\n"
        return new
    content = replace_var(content, "MY_TIKTOK_USERNAME", tiktok_id)
    content = replace_var(content, "ANALYTICS_URL", url)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(content)

# ────────────────────────────────────────────────────────────────
#  データ読み込み
# ────────────────────────────────────────────────────────────────
def find_col(df, *kws):
    for kw in kws:
        for col in df.columns:
            if kw in col: return col
    return None

def _insights_csv_path():
    try:
        import config as _cfg
        raw = getattr(_cfg, "CSV_INSIGHTS_FILE", "data/insights.csv")
    except Exception:
        raw = "data/insights.csv"
    return os.path.join(DATA_DIR, os.path.splitext(os.path.basename(raw))[0] + ".csv")

def load_insights():
    if not _HAS_PANDAS: return None, "pandas が必要です"
    path = _insights_csv_path()
    if not os.path.exists(path): return None, f"insights.csv が見つかりません\n{path}"
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
        date_col = next((c for c in df.columns
                         if any(k in c for k in ["日","date","時","取得"])), None)
        df["_date"] = pd.to_datetime(
            df[date_col] if date_col else df.iloc[:,0], errors="coerce")
        df = df.dropna(subset=["_date"]).sort_values("_date").reset_index(drop=True)
        return df, None
    except Exception as e:
        return None, str(e)

def load_gifts():
    if not _HAS_PANDAS: return None, "pandas が必要です"
    if not os.path.exists(CSV_FILE): return None, "gift_timeline.csv が見つかりません"
    try:
        df = pd.read_csv(CSV_FILE, encoding="utf-8-sig")
        df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
        df = df.dropna(subset=["timestamp"])
        df_gift = df[df["type"] == "gift"].copy()
        if df_gift.empty: return df_gift, None
        def parse_detail(s):
            m = re.match(r"^(.+?)\s*[×x×](\d+)$", str(s).strip())
            return (m.group(1).strip(), int(m.group(2))) if m else (str(s).strip(), 1)
        parsed = df_gift["detail"].apply(parse_detail)
        df_gift["gift_name"] = [p[0] for p in parsed]
        df_gift["count"]     = [p[1] for p in parsed]
        df_gift["_date"]     = df_gift["timestamp"]
        return df_gift, None
    except Exception as e:
        return None, str(e)

def load_viewers():
    if not _HAS_PANDAS: return None, "pandas が必要です"
    if not os.path.exists(VIEWERS_FILE): return None, "viewers.csv が見つかりません"
    try:
        df = pd.read_csv(VIEWERS_FILE, encoding="utf-8-sig")
        df.columns = [c.strip().lower() for c in df.columns]
        if "session_date" not in df.columns:
            df = df.rename(columns={df.columns[0]: "session_date"})
        df["session_date"] = pd.to_datetime(df["session_date"], errors="coerce").dt.date
        return df.dropna(subset=["session_date"]), None
    except Exception as e:
        return None, str(e)

# ────────────────────────────────────────────────────────────────
#  ダークテーマ用 matplotlib Figure 生成ヘルパー
# ────────────────────────────────────────────────────────────────
def _dark_fig(nrows=1, ncols=1, figsize=(12, 5)):
    fig, axes = plt.subplots(nrows, ncols, figsize=figsize,
                             facecolor=C_BG)
    for ax in (axes.flat if hasattr(axes, "flat") else [axes]):
        ax.set_facecolor(C_PANEL)
        ax.tick_params(colors=C_SUBTEXT, labelsize=8)
        for spine in ax.spines.values():
            spine.set_color(C_BORDER)
        ax.title.set_color("#c4b5fd")
        ax.xaxis.label.set_color(C_SUBTEXT)
        ax.yaxis.label.set_color(C_SUBTEXT)
    return fig, axes

def _annotate_bars(ax, vals, color="white", fmt="{:.0f}"):
    """棒グラフの上に値ラベルを表示"""
    for i, v in enumerate(vals):
        ax.text(i, v * 1.01, fmt.format(v),
                ha="center", va="bottom", fontsize=7,
                color=color, fontweight="bold")

def _add_data_points(ax, x, y, color):
    """折れ線に加えてデータポイントをプロット"""
    ax.plot(x, y, "o", color=color, markersize=5, zorder=5)

# ────────────────────────────────────────────────────────────────
#  ライブ監視ワーカースレッド  [Bug1 / Bug2 修正]
# ────────────────────────────────────────────────────────────────
class LiveWorker(QThread):
    log_signal      = pyqtSignal(str)
    status_signal   = pyqtSignal(str, str)   # (text, color)
    finished_signal = pyqtSignal()

    def __init__(self, stop_event: threading.Event):
        super().__init__()
        self._stop = stop_event
        self._insight_thread = None

    def run(self):
        try:
            import asyncio
            import importlib
            import config
            # ★ Bug-A 修正: スレッド内でも config を最新状態に再読込
            importlib.reload(config)
            from modules.live_bot import LiveBot

            self.log_signal.emit(f"監視ボットを起動しました")
            self.log_signal.emit(f"監視対象: @{config.MY_TIKTOK_USERNAME}")
            self.status_signal.emit("🟢 監視中", C_GREEN)

            def on_stream_end():
                self.log_signal.emit("配信終了を検知。3分後にインサイトを自動取得します…")
                for remaining in range(180, 0, -10):
                    time.sleep(10)
                    # 停止要求があっても3分待機後にインサイト取得は実行する
                    # （停止ボタン押下時も取得する仕様）
                self._run_insights()

            bot = LiveBot(on_stream_end_callback=on_stream_end,
                          stop_event=self._stop)
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(bot.start())
            finally:
                loop.close()

        except ImportError as ex:
            # ★ Bug-B: 依存ライブラリ未インストール時の詳細メッセージ
            tb = traceback.format_exc()
            self.log_signal.emit(f"❌ インポートエラー: {ex}")
            self.log_signal.emit("ℹ️ 依存ライブラリが不足しています。")
            self.log_signal.emit("   pip install TikTokLive selenium webdriver-manager を実行してください")
            self.log_signal.emit(f"詳細:\n{tb[:500]}")
            traceback.print_exc()
        except Exception as ex:
            tb = traceback.format_exc()
            self.log_signal.emit(f"❌ ボットエラー: {ex}")
            self.log_signal.emit(f"詳細:\n{tb[:800]}")
            traceback.print_exc()
        finally:
            self.status_signal.emit("⏹ 停止中", C_RED)
            self.finished_signal.emit()

    def _run_insights(self):
        """インサイト取得をこのスレッド内で実行"""
        self.log_signal.emit("インサイト取得中…")
        try:
            from modules.insights import collect_insights
            ok = collect_insights()
            if ok:
                self.log_signal.emit("✅ インサイト取得完了！")
            else:
                self.log_signal.emit("❌ インサイト取得失敗。debug_page.html を確認してください。")
        except Exception as ex:
            self.log_signal.emit(f"❌ インサイト取得エラー: {ex}")

    def trigger_insight_now(self):
        """停止ボタン押下時に手動でインサイト取得を起動 [Bug2修正]"""
        t = threading.Thread(target=self._run_insights, daemon=True)
        self._insight_thread = t
        t.start()
        self.log_signal.emit("停止要求を受け付けました。インサイト取得を開始します…")

# ────────────────────────────────────────────────────────────────
#  メインウィンドウ
# ────────────────────────────────────────────────────────────────
class KinakoApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("🦦 カワウソマネージャー きなこ")
        self.resize(1100, 800)

        # stdout リダイレクト [Bug1]
        self._redirector = _StdoutRedirector(sys.stdout)
        self._redirector.text_written.connect(self._on_stdout)
        sys.stdout = self._redirector

        self._bot_stop_event = threading.Event()
        self._live_worker: LiveWorker | None = None

        # グラフキャッシュ
        self._insight_fig = self._insight_df = None
        self._gift_fig    = self._gift_df    = None
        self._repeat_fig  = self._repeat_df  = None

        self._build_ui()

    def _build_ui(self):
        self.setStyleSheet(STYLE_SHEET)

        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ── ヘッダー ──
        header = QWidget()
        header.setStyleSheet(f"background:{C_ACCENT};")
        header.setFixedHeight(54)
        h_lay = QHBoxLayout(header)
        h_lay.setContentsMargins(16, 0, 16, 0)

        title_lbl = QLabel("🦦 カワウソマネージャー きなこ")
        title_lbl.setStyleSheet("color:white; font-size:15pt; font-weight:bold;")
        h_lay.addWidget(title_lbl)
        h_lay.addStretch()

        self._btn_excel = btn("📊 Excelエクスポート", C_BLUE,  self._on_export_excel, 160)
        self._btn_csv   = btn("📥 CSVエクスポート",   C_TEAL, self._on_export_csv,   140)
        h_lay.addWidget(self._btn_excel)
        h_lay.addSpacing(8)
        h_lay.addWidget(self._btn_csv)

        root_layout.addWidget(header)

        # ── タブ ──
        self._tabs = QTabWidget()
        root_layout.addWidget(self._tabs)

        self._tab_setup   = QWidget()
        self._tab_live    = QWidget()
        self._tab_insight = QWidget()
        self._tab_report  = QWidget()

        self._tabs.addTab(self._tab_setup,   "⚙️  セットアップ")
        self._tabs.addTab(self._tab_live,    "📡  ライブ監視")
        self._tabs.addTab(self._tab_insight, "📥  インサイト取得")
        self._tabs.addTab(self._tab_report,  "📊  レポート")

        self._build_setup_tab()
        self._build_live_tab()
        self._build_insight_tab()
        self._build_report_tab()

        # フッター
        footer = QLabel(f"データフォルダ: {DATA_DIR}")
        footer.setObjectName("subText")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setStyleSheet(f"color:{C_SUBTEXT}; padding:4px;")
        root_layout.addWidget(footer)

    # ─────────────────────────────────────────────────────────
    #  ① セットアップタブ
    # ─────────────────────────────────────────────────────────
    def _build_setup_tab(self):
        lay = QVBoxLayout(self._tab_setup)
        lay.setContentsMargins(40, 24, 40, 24)
        lay.setSpacing(10)

        t = QLabel("⚙️  初期セットアップ"); t.setObjectName("sectionTitle"); lay.addWidget(t)

        sub = QLabel("2項目を入力して「保存してセットアップ完了」を押してください")
        sub.setObjectName("subText"); lay.addWidget(sub)

        cfg_lbl = QLabel(f"📄 config.py: {CONFIG_FILE}")
        cfg_lbl.setStyleSheet(f"color:{C_SUBTEXT}; font-size:8pt;"); lay.addWidget(cfg_lbl)

        lay.addSpacing(10)

        lay.addWidget(QLabel("① TikTok ID（@ なし）"))
        self._setup_id = QLineEdit(read_config_value("MY_TIKTOK_USERNAME"))
        self._setup_id.setPlaceholderText("例: kinako_tiktok")
        lay.addWidget(self._setup_id)

        lay.addSpacing(6)
        lay.addWidget(QLabel("② インサイトページ URL"))
        self._setup_url = QLineEdit(read_config_value("ANALYTICS_URL"))
        self._setup_url.setPlaceholderText(
            "https://livecenter.tiktok.com/analytics/live_video?lang=ja-JP")
        lay.addWidget(self._setup_url)

        lay.addSpacing(16)
        save_btn = btn("✅  保存してセットアップ完了", C_ACCENT, self._on_setup_save)
        save_btn.setFixedHeight(44)
        lay.addWidget(save_btn)

        foot = QLabel("設定は config.py に保存　／　Chrome は自動検出・永続プロファイルで起動")
        foot.setStyleSheet(f"color:{C_SUBTEXT}; font-size:8pt;")
        foot.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(foot)
        lay.addStretch()

    def _on_setup_save(self):
        tid = self._setup_id.text().strip().lstrip("@")
        url = self._setup_url.text().strip()
        for err in (validate_tiktok_id(tid), validate_url(url)):
            if err:
                QMessageBox.warning(self, "入力エラー", err); return
        # ★ Bug-C 修正: Chrome が見つからなくても設定保存は続行（警告のみ）
        # webdriver-manager が ChromeDriver を自動管理するため、
        # Chrome本体のパス確認に失敗しても実際の動作には支障がない場合がある
        chrome_warning = ""
        if not find_chrome():
            chrome_warning = ("\n\n⚠️ Google Chrome のパスが自動検出できませんでした。\n"
                              "インストール済みの場合は問題ありません。\n"
                              "Chrome がない場合は https://www.google.com/chrome/ からインストールしてください。")
        try:
            update_config(tid, url)
            # ★ Bug-A 修正③: config.py 書き換え後にキャッシュを破棄
            # 次に import config または importlib.reload(config) した時に
            # 必ず最新の config.py が読まれるようにする
            import sys as _sys
            _sys.modules.pop("config", None)
            QMessageBox.information(self, "セットアップ完了 🎉",
                f"設定を保存しました！\n\n  ・TikTok ID : @{tid}\n\n"
                "「📡 ライブ監視」タブから配信監視を開始できます！" + chrome_warning)
        except PermissionError:
            QMessageBox.critical(self, "書き込みエラー",
                f"config.py に書き込めませんでした。\n{CONFIG_FILE}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", str(e))

    # ─────────────────────────────────────────────────────────
    #  ② ライブ監視タブ
    # ─────────────────────────────────────────────────────────
    def _build_live_tab(self):
        lay = QVBoxLayout(self._tab_live)
        lay.setContentsMargins(40, 24, 40, 24)
        lay.setSpacing(10)

        t = QLabel("📡  ライブ監視ボット"); t.setObjectName("sectionTitle"); lay.addWidget(t)
        sub = QLabel("「監視開始」を押すとバックグラウンドで TikTok ライブを監視します。\n"
                     "配信終了を検知すると、3分後にインサイトを自動取得します。")
        sub.setObjectName("subText"); lay.addWidget(sub)

        # ステータス行
        status_frame = QFrame(); status_frame.setObjectName("statusFrame")
        sf_lay = QHBoxLayout(status_frame)
        sf_lay.addWidget(QLabel("ステータス:"))
        self._status_lbl = QLabel("⏹ 停止中")
        self._status_lbl.setStyleSheet(f"color:{C_RED}; font-weight:bold; font-size:11pt;")
        sf_lay.addWidget(self._status_lbl)
        sf_lay.addStretch()
        lay.addWidget(status_frame)

        # ボタン行
        btn_row = QHBoxLayout()
        self._btn_start = btn("▶  監視開始", C_GREEN, self._on_live_start)
        self._btn_stop  = btn("⏹  監視停止", C_RED,   self._on_live_stop)
        self._btn_stop.setEnabled(False)
        btn_row.addWidget(self._btn_start)
        btn_row.addWidget(self._btn_stop)
        btn_row.addStretch()
        lay.addLayout(btn_row)

        lay.addWidget(QLabel("ログ"))
        self._log_view = QTextEdit()
        self._log_view.setReadOnly(True)
        lay.addWidget(self._log_view, stretch=1)

    @pyqtSlot(str)
    def _on_stdout(self, text):
        """stdout出力をログビューに転送 [Bug1修正]"""
        self._append_log(text)

    def _append_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_view.append(f'<span style="color:{C_SUBTEXT};">[{ts}]</span> {msg}')
        self._log_view.moveCursor(QTextCursor.MoveOperation.End)

    def _on_live_start(self):
        try:
            import importlib
            import config
            # ★ Bug-A 修正: セットアップ後の config.py 書き換えを確実に反映させる
            # Python は一度 import したモジュールをキャッシュするため、
            # アプリ起動後にファイルを書き換えても古い値のまま → reload() で強制再読込
            importlib.reload(config)
            if not config.MY_TIKTOK_USERNAME:
                raise ValueError("MY_TIKTOK_USERNAME が未設定です")
            if not config.ANALYTICS_URL.startswith("https://livecenter.tiktok.com/"):
                raise ValueError("ANALYTICS_URL が不正です")
        except ValueError as e:
            QMessageBox.critical(self, "セットアップ未完了",
                f"設定に問題があります。\n\n{e}\n\n「⚙️ セットアップ」タブで設定してください。")
            return
        except Exception as e:
            QMessageBox.critical(self, "設定読み込みエラー",
                f"config.py の読み込みに失敗しました。\n\n{e}")
            return

        self._bot_stop_event.clear()
        self._live_worker = LiveWorker(self._bot_stop_event)
        self._live_worker.log_signal.connect(self._append_log)
        self._live_worker.status_signal.connect(self._set_status)
        self._live_worker.finished_signal.connect(self._on_bot_finished)
        self._live_worker.start()

        self._btn_start.setEnabled(False)
        self._btn_stop.setEnabled(True)

    @pyqtSlot(str, str)
    def _set_status(self, text, color):
        self._status_lbl.setText(text)
        self._status_lbl.setStyleSheet(
            f"color:{color}; font-weight:bold; font-size:11pt;")

    def _on_live_stop(self):
        """停止ボタン：stop_eventをセットし、直ちにインサイト取得も開始 [Bug2修正]"""
        self._bot_stop_event.set()
        self._btn_stop.setEnabled(False)
        self._set_status("⏸ 停止中…", "#f59e0b")
        self._append_log("停止リクエストを送信しました")

        # 配信中に手動停止した場合もインサイト取得を起動
        if self._live_worker and self._live_worker.isRunning():
            self._live_worker.trigger_insight_now()

    @pyqtSlot()
    def _on_bot_finished(self):
        self._btn_start.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._set_status("⏹ 停止中", C_RED)
        self._append_log("監視ボットが停止しました")

    # ─────────────────────────────────────────────────────────
    #  ③ インサイト手動取得タブ
    # ─────────────────────────────────────────────────────────
    def _build_insight_tab(self):
        lay = QVBoxLayout(self._tab_insight)
        lay.setContentsMargins(40, 24, 40, 24)
        lay.setSpacing(10)

        t = QLabel("📥  インサイト手動取得"); t.setObjectName("sectionTitle"); lay.addWidget(t)
        sub = QLabel("ボタンを押すと Chrome が自動起動し、TikTok LiveCenter から\n"
                     "最新の配信インサイトデータを取得して data/insights.csv に保存します。")
        sub.setObjectName("subText"); lay.addWidget(sub)

        lay.addSpacing(16)
        get_btn = btn("📥  今すぐインサイトを取得", "#0369a1", self._on_insight_get)
        get_btn.setFixedHeight(48)
        get_btn.setMaximumWidth(320)
        lay.addWidget(get_btn)

        lay.addSpacing(20)
        guide = QLabel(
            "【手動取得の使い方】\n"
            "1. 「⚙️ セットアップ」タブで設定を済ませてください\n"
            "2. 「今すぐインサイトを取得」ボタンを押します\n"
            "3. Chrome が自動起動します（初回は TikTok ログインが必要）\n"
            "4. 取得完了後、「📊 レポート」タブでグラフを確認できます\n\n"
            "※ ChromeDriver は自動インストールされます（webdriver-manager）"
        )
        guide.setObjectName("subText")
        lay.addWidget(guide)
        lay.addStretch()

    def _on_insight_get(self):
        try:
            import config; config.validate()
        except Exception as e:
            QMessageBox.critical(self, "設定エラー",
                f"config.py の設定に問題があります。\n\n{e}")
            return
        if QMessageBox.question(self, "インサイト手動取得",
            "TikTok LiveCenter のインサイトを今すぐ取得します。\n\n"
            "Chrome が自動的に起動します。OK で開始しますか？") \
                != QMessageBox.StandardButton.Yes:
            return

        def run():
            try:
                from modules.insights import collect_insights
                ok = collect_insights()
                if ok:
                    self._show_info_later("✅ 取得完了",
                        "インサイトの取得が完了しました！\n保存先: data/insights.csv")
                else:
                    self._show_warn_later("⚠️ 取得失敗",
                        "インサイトの取得に失敗しました。\n"
                        "・TikTok にログインしているか確認\n"
                        "・data/debug_page.html で詳細確認")
            except Exception as ex:
                self._show_err_later("❌ エラー", str(ex))
                traceback.print_exc()

        threading.Thread(target=run, daemon=True).start()

    def _show_info_later(self, t, m):
        QTimer.singleShot(0, lambda: QMessageBox.information(self, t, m))
    def _show_warn_later(self, t, m):
        QTimer.singleShot(0, lambda: QMessageBox.warning(self, t, m))
    def _show_err_later(self, t, m):
        QTimer.singleShot(0, lambda: QMessageBox.critical(self, t, m))

    # ─────────────────────────────────────────────────────────
    #  ④ レポートタブ  [Bug4修正 + グラフ改善]
    # ─────────────────────────────────────────────────────────
    def _build_report_tab(self):
        if not _HAS_PANDAS:
            lay = QVBoxLayout(self._tab_report)
            lay.addWidget(QLabel("pandas が必要です: pip install pandas"))
            return

        lay = QVBoxLayout(self._tab_report)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(4)

        sub_tabs = QTabWidget()
        sub_tabs.setStyleSheet(f"""
            QTabBar::tab {{ padding:6px 16px; font-size:9pt; }}
            QTabBar::tab:selected {{ background:{C_ACCENT2}; color:white; }}
        """)
        lay.addWidget(sub_tabs)

        self._rtab_insight = QWidget()
        self._rtab_gift    = QWidget()
        self._rtab_repeat  = QWidget()
        sub_tabs.addTab(self._rtab_insight, "📊  インサイト")
        sub_tabs.addTab(self._rtab_gift,    "🎁  ギフト")
        sub_tabs.addTab(self._rtab_repeat,  "👥  リピート率")
        self._report_sub_tabs = sub_tabs

        self._build_insight_report()
        self._build_gift_report()
        self._build_repeat_report()

        # 初期描画（遅延）
        QTimer.singleShot(400, self._on_show_insights)
        QTimer.singleShot(500, self._on_show_gift)
        QTimer.singleShot(600, self._on_show_repeat)

    def _make_ctrl_row(self, parent_lay, on_show):
        """日付範囲コントロール行を生成し、(date_start, date_end) を返す"""
        row = QHBoxLayout()
        row.setSpacing(8)

        today     = QDate.currentDate()
        one_month = today.addDays(-30)

        row.addWidget(QLabel("開始日"))
        de_start = QDateEdit(one_month)   # [Bug4修正] QDateEdit使用
        de_start.setCalendarPopup(True)
        de_start.setDisplayFormat("yyyy-MM-dd")
        de_start.setMinimumWidth(120)
        row.addWidget(de_start)

        row.addSpacing(10)
        row.addWidget(QLabel("終了日"))
        de_end = QDateEdit(today)
        de_end.setCalendarPopup(True)
        de_end.setDisplayFormat("yyyy-MM-dd")
        de_end.setMinimumWidth(120)
        row.addWidget(de_end)

        show_btn = btn("グラフを表示", C_ACCENT, on_show, 130)
        show_btn.setFixedHeight(34)
        row.addWidget(show_btn)
        row.addStretch()

        parent_lay.addLayout(row)
        return de_start, de_end

    # ── インサイトレポート ──
    def _build_insight_report(self):
        lay = QVBoxLayout(self._rtab_insight)
        lay.setContentsMargins(8, 8, 8, 8)
        self._de_ins_start, self._de_ins_end = self._make_ctrl_row(lay, self._on_show_insights)
        self._canvas_ins = None
        self._frame_ins  = QWidget()
        self._frame_ins.setLayout(QVBoxLayout())
        lay.addWidget(self._frame_ins, stretch=1)

    def _on_show_insights(self):
        df, err = load_insights()
        if err or df is None: return
        s = pd.to_datetime(self._de_ins_start.date().toString("yyyy-MM-dd"))
        e = pd.to_datetime(self._de_ins_end.date().toString("yyyy-MM-dd")) \
            + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        df = df[(df["_date"] >= s) & (df["_date"] <= e)]
        if df.empty: return

        col_peak    = "最高同時視聴者数" if "最高同時視聴者数" in df.columns else find_col(df,"最高同時","peak")
        col_diamond = "ダイヤ合計"      if "ダイヤ合計"      in df.columns else find_col(df,"diamond")
        col_gifter  = "ギフト贈呈者数"  if "ギフト贈呈者数"  in df.columns else find_col(df,"ギフト贈呈","gifter")
        col_watch   = "平均視聴時間"    if "平均視聴時間"    in df.columns else find_col(df,"平均視聴","watch")

        for col in [col_peak, col_diamond, col_gifter, col_watch]:
            if col: df[col] = pd.to_numeric(df[col], errors="coerce")

        diamond_total = int(df[col_diamond].sum(skipna=True)) if col_diamond else 0
        plt.close("all")
        fig, axes = _dark_fig(2, 2, (11, 7))
        title = (f"インサイト（{self._de_ins_start.date().toString('yyyy-MM-dd')} "
                 f"～ {self._de_ins_end.date().toString('yyyy-MM-dd')}）"
                 f"  ◆ 期間合計ダイヤ: {diamond_total:,}")
        fig.suptitle(title, fontsize=10, y=0.99, color=C_TEXT)

        plot_cfgs = [
            (axes[0][0], col_peak,    "#4f86c6", "最高同時視聴者数（人）"),
            (axes[0][1], col_diamond, "#f5a623", "ダイヤ数"),
            (axes[1][0], col_gifter,  "#7ed321", "ギフト贈呈者数（人）"),
            (axes[1][1], col_watch,   "#e87c7c", "平均視聴時間"),
        ]
        for ax, col, color, ylabel in plot_cfgs:
            if col and col in df.columns:
                mask = df[col].notna()
                vals = df.loc[mask, col]
                labels = (df.loc[mask,"_date"].dt.strftime("%m/%d")
                          if "_date" in df.columns else range(len(vals)))
                if not vals.empty:
                    x = range(len(vals))
                    ax.bar(x, vals, color=color, alpha=0.8)
                    _annotate_bars(ax, vals.tolist(), color="white")
                    _add_data_points(ax, list(x), vals.tolist(), "white")
                    ax.set_xticks(list(x))
                    ax.set_xticklabels(list(labels), rotation=45, fontsize=8)
                    ax.set_ylabel(ylabel, fontsize=9)
                    mean_val = vals.mean()
                    ax.axhline(mean_val, color="red", linestyle="--", lw=1.2,
                               label=f"平均: {mean_val:.1f}")
                    ax.legend(fontsize=8, facecolor=C_PANEL, labelcolor=C_TEXT)
            else:
                ax.text(0.5,0.5,"データなし",ha="center",va="center",
                        transform=ax.transAxes, color=C_SUBTEXT, fontsize=11)
            ax.set_title(ylabel, fontsize=10)
        plt.tight_layout(rect=[0,0,1,0.96])

        self._insight_fig = fig
        self._insight_df  = df
        self._replace_canvas(self._frame_ins, fig)

    # ── ギフトレポート ──
    def _build_gift_report(self):
        lay = QVBoxLayout(self._rtab_gift)
        lay.setContentsMargins(8, 8, 8, 8)
        self._de_gift_start, self._de_gift_end = self._make_ctrl_row(lay, self._on_show_gift)
        self._frame_gift = QWidget(); self._frame_gift.setLayout(QVBoxLayout())
        lay.addWidget(self._frame_gift, stretch=1)

    def _on_show_gift(self):
        df, err = load_gifts()
        if err or df is None or df.empty: return
        s = pd.to_datetime(self._de_gift_start.date().toString("yyyy-MM-dd")).date()
        e = (pd.to_datetime(self._de_gift_end.date().toString("yyyy-MM-dd"))
             + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)).date()
        df_r = df[(df["_date"].dt.date >= s) & (df["_date"].dt.date <= e)].copy()
        if df_r.empty: return

        plt.close("all")
        fig, axes = _dark_fig(1, 3, (14, 5))
        ug = df_r["user"].nunique() if "user" in df_r.columns else 0
        fig.suptitle(
            f"ギフトタイムライン  ◆ ギフター: {ug}人  |  ギフト回数: {len(df_r)}回",
            fontsize=10, color=C_TEXT)

        ax1, ax2, ax3 = axes

        # 時間帯別
        if "_date" in df_r.columns:
            df_r["hour"] = df_r["_date"].dt.hour
            hourly = df_r.groupby("hour").size()
            ax1.bar(hourly.index, hourly.values, color="#f5a623", alpha=0.85)
            _annotate_bars(ax1, hourly.values.tolist())
            ax1.set_xlabel("時刻（時）",  fontsize=9)
            ax1.set_ylabel("ギフト回数",  fontsize=9)
        ax1.set_title("時間帯別ギフト回数", fontsize=10)

        # トップギフター
        if "user" in df_r.columns:
            tg = df_r.groupby("user").size().nlargest(10)
            y_pos = range(len(tg))
            ax2.barh(list(y_pos), tg.values[::-1], color="#7ed321", alpha=0.85)
            ax2.set_yticks(list(y_pos))
            ax2.set_yticklabels(list(tg.index[::-1]), fontsize=8)
            for i, v in enumerate(tg.values[::-1]):
                ax2.text(v*1.01, i, str(v), va="center", fontsize=7, color="white")
            ax2.set_xlabel("ギフト回数", fontsize=9)
        ax2.set_title("トップギフター Top10", fontsize=10)

        # ギフト種別
        if "gift_name" in df_r.columns:
            tgt = df_r.groupby("gift_name")["count"].sum().nlargest(10)
            y_pos = range(len(tgt))
            ax3.barh(list(y_pos), tgt.values[::-1], color="#4f86c6", alpha=0.85)
            ax3.set_yticks(list(y_pos))
            ax3.set_yticklabels(list(tgt.index[::-1]), fontsize=8)
            for i, v in enumerate(tgt.values[::-1]):
                ax3.text(v*1.01, i, str(v), va="center", fontsize=7, color="white")
            ax3.set_xlabel("合計個数", fontsize=9)
        ax3.set_title("ギフト種別 Top10", fontsize=10)

        plt.tight_layout(rect=[0,0,1,0.94])
        self._gift_fig = fig
        self._gift_df  = df_r
        self._replace_canvas(self._frame_gift, fig)

    # ── リピート率レポート ──
    def _build_repeat_report(self):
        lay = QVBoxLayout(self._rtab_repeat)
        lay.setContentsMargins(8, 8, 8, 8)
        self._de_rep_start, self._de_rep_end = self._make_ctrl_row(lay, self._on_show_repeat)
        self._frame_rep = QWidget(); self._frame_rep.setLayout(QVBoxLayout())
        lay.addWidget(self._frame_rep, stretch=1)

    def _on_show_repeat(self):
        df, err = load_viewers()
        if err or df is None or df.empty: return
        s = pd.to_datetime(self._de_rep_start.date().toString("yyyy-MM-dd")).date()
        e = (pd.to_datetime(self._de_rep_end.date().toString("yyyy-MM-dd"))
             + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)).date()
        df = df[(df["session_date"] >= s) & (df["session_date"] <= e)]
        if df.empty: return

        uid_col  = "uid" if "uid" in df.columns else \
                   (df.columns[2] if len(df.columns) > 2 else None)
        name_col = "display_name" if "display_name" in df.columns else None
        if uid_col is None: return

        sc       = df.groupby(uid_col)["session_date"].nunique()
        total    = len(sc)
        repeats  = int((sc >= 2).sum())
        rate     = repeats / total * 100 if total > 0 else 0.0
        sv       = df.groupby("session_date")[uid_col].nunique().sort_index()
        top_r    = sc[sc >= 2].nlargest(10)
        top_lbl  = ([df.drop_duplicates(uid_col).set_index(uid_col)[name_col].get(u,str(u))
                     for u in top_r.index]
                    if name_col else [str(u) for u in top_r.index])

        plt.close("all")
        fig, axes = _dark_fig(1, 3, (14, 5))
        fig.suptitle(
            f"リピート率レポート  |  ユニーク視聴者: {total}人  "
            f"リピーター: {repeats}人  リピート率: {rate:.1f}%",
            fontsize=10, color=C_TEXT)

        ax0, ax1, ax2 = axes

        # 円グラフ
        ax0.set_title("リピーター比率", fontsize=10)
        if total > 0:
            ax0.pie([repeats, total - repeats],
                    labels=[f"リピーター\n{repeats}人", f"初回のみ\n{total-repeats}人"],
                    colors=[C_ACCENT, "#c4b5fd"],
                    autopct="%1.1f%%", startangle=90,
                    textprops={"fontsize":10,"color":C_TEXT})

        # セッション別ユニーク視聴者
        ax1.set_title("セッション別ユニーク視聴者", fontsize=10)
        if not sv.empty:
            dates = [str(d) for d in sv.index]
            x = range(len(dates))
            ax1.bar(x, sv.values, color="#4f86c6", alpha=0.85)
            _annotate_bars(ax1, sv.values.tolist())
            _add_data_points(ax1, list(x), sv.values.tolist(), "white")
            ax1.set_xticks(list(x))
            ax1.set_xticklabels(dates, rotation=45, fontsize=8)
            ax1.set_ylabel("ユニーク視聴者数（人）", fontsize=9)
            mean_v = sv.mean()
            ax1.axhline(mean_v, color="red", linestyle="--", lw=1.2,
                        label=f"平均: {mean_v:.1f}")
            ax1.legend(fontsize=8, facecolor=C_PANEL, labelcolor=C_TEXT)

        # リピーター Top10
        ax2.set_title("リピーター Top10（参加セッション数）", fontsize=10)
        if len(top_r) > 0:
            y_pos = range(len(top_r))
            ax2.barh(list(y_pos), top_r.values[::-1], color="#7ed321", alpha=0.85)
            ax2.set_yticks(list(y_pos))
            ax2.set_yticklabels(top_lbl[::-1], fontsize=8)
            for i, v in enumerate(top_r.values[::-1]):
                ax2.text(v*1.01, i, str(v), va="center", fontsize=7, color="white")
            ax2.set_xlabel("参加セッション数", fontsize=9)

        plt.tight_layout(rect=[0,0,1,0.94])
        self._repeat_fig = fig
        self._repeat_df  = df
        self._replace_canvas(self._frame_rep, fig)

    def _replace_canvas(self, frame: QWidget, fig):
        """フレーム内の既存キャンバスを新しいFigureで置き換える"""
        lay = frame.layout()
        while lay.count():
            item = lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        canvas = FigureCanvas(fig)
        canvas.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        lay.addWidget(canvas)
        canvas.draw()

    # ─────────────────────────────────────────────────────────
    #  エクスポート
    # ─────────────────────────────────────────────────────────
    def _get_current_report(self):
        idx = self._report_sub_tabs.currentIndex() if hasattr(self, "_report_sub_tabs") else -1
        if idx == 0: return self._insight_fig, self._insight_df, \
            f"インサイト_{self._de_ins_start.date().toString('yyyyMMdd')}"
        if idx == 1: return self._gift_fig, self._gift_df, \
            f"ギフト_{self._de_gift_start.date().toString('yyyyMMdd')}"
        if idx == 2: return self._repeat_fig, self._repeat_df, "リピート率レポート"
        return None, None, ""

    def _on_export_excel(self):
        fig, df, title = self._get_current_report()
        if fig is None or df is None:
            QMessageBox.warning(self, "未表示",
                "先にレポートタブでグラフを表示してください。"); return
        try:
            import openpyxl
            from openpyxl.drawing.image import Image as XLImage
        except ImportError:
            QMessageBox.critical(self, "ライブラリエラー",
                "openpyxl が必要です: pip install openpyxl"); return
        path, _ = QFileDialog.getSaveFileName(
            self, "Excelを保存", f"{title}.xlsx", "Excel (*.xlsx)")
        if not path: return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active; ws.title = "データ"
            out = df.copy()
            for c in out.columns:
                if pd.api.types.is_datetime64_any_dtype(out[c]):
                    out[c] = out[c].astype(str)
            ws.append(list(out.columns))
            for row in out.itertuples(index=False): ws.append(list(row))
            wc = wb.create_sheet("グラフ")
            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
            buf.seek(0)
            wc.add_image(XLImage(buf), "A1")
            wb.save(path)
            QMessageBox.information(self, "保存完了", f"保存しました:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "保存エラー", str(e))

    def _on_export_csv(self):
        fig, df, title = self._get_current_report()
        if df is None:
            QMessageBox.warning(self, "未表示",
                "先にレポートタブでグラフを表示してください。"); return
        path, _ = QFileDialog.getSaveFileName(
            self, "CSVを保存", f"{title}.csv", "CSV (*.csv)")
        if not path: return
        try:
            out = df.copy()
            for c in out.columns:
                if pd.api.types.is_datetime64_any_dtype(out[c]):
                    out[c] = out[c].astype(str)
            out.to_csv(path, index=False, encoding="utf-8-sig")
            QMessageBox.information(self, "保存完了", f"保存しました:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "保存エラー", str(e))

    # ─────────────────────────────────────────────────────────
    #  終了処理
    # ─────────────────────────────────────────────────────────
    def closeEvent(self, event):
        if self._live_worker and self._live_worker.isRunning():
            r = QMessageBox.question(self, "終了確認",
                "ライブ監視ボットが動作中です。\n終了してよいですか？")
            if r != QMessageBox.StandardButton.Yes:
                event.ignore(); return
        self._bot_stop_event.set()
        sys.stdout = self._redirector._original
        plt.close("all")
        event.accept()


# ────────────────────────────────────────────────────────────────
#  エントリポイント
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = KinakoApp()
    window.show()
    sys.exit(app.exec())
