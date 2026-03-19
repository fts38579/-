# ===================================================
#  config.py  –  カワウソマネージャー きなこ 設定ファイル
# ===================================================
#
#  ⚠️ このファイルを直接編集しないでください。
#     「セットアップ」フォルダの「初期セットアップ.py」から設定してください。
#
# ===================================================
import os

# ---------- TikTok ----------
# ※ 初期セットアップ.py によって自動で書き換えられます
MY_TIKTOK_USERNAME = ''           # セットアップ前は空文字
ANALYTICS_URL      = ''

# ---------- データファイル ----------
CSV_FILE           = 'data/gift_timeline.csv'
CSV_INSIGHTS_FILE  = 'data/insights.csv'   # 旧: EXCEL_FILE → リネーム
VIEWERS_FILE       = 'data/viewers.csv'
COMMENTS_FILE      = 'data/comments.csv'

# ---------- Chrome 永続プロファイル ----------
# ショートカット（.lnk）は不要。このフォルダにTikTokのログイン情報が保持されます。
CHROME_PROFILE_DIR = os.path.join(
    os.environ.get('LOCALAPPDATA', os.path.expanduser('~')),
    'TikTokChromeProfile'
)

# ===================================================
#  起動前バリデーション（セットアップ未実施の早期検知）
# ===================================================
def _validate():
    """
    必須項目が未設定の場合は起動を止めて案内メッセージを出す。
    main.py / report_gui.py の冒頭で config.validate() として呼び出す。
    """
    errors = []
    if not MY_TIKTOK_USERNAME:
        errors.append("・MY_TIKTOK_USERNAME が未設定です")
    if not ANALYTICS_URL.startswith("https://livecenter.tiktok.com/"):
        errors.append("・ANALYTICS_URL が不正です")
    if errors:
        import tkinter as tk
        from tkinter import messagebox
        tk.Tk().withdraw()
        messagebox.showerror(
            "セットアップが未完了です",
            "以下の設定が未完了のため起動できません。\n\n"
            + "\n".join(errors)
            + "\n\n「セットアップ」フォルダ内の\n"
              "「初期セットアップ.py」を実行してください。"
        )
        raise SystemExit(1)

validate = _validate  # 外部から config.validate() で呼び出せるようにする

# ── TikTok LiveCenter – 最新行クリック用 XPath ──────────────────────
# ※ TikTok の HTML 構造変更時は更新が必要
XPATH_TOP_ROW = '(//tr[contains(@class,"tt-live-table-row")])[1]'
