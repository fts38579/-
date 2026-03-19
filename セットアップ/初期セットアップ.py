# 初期セットアップ.py
# ──────────────────────────────────────────────────────────
# このファイルは「セットアップ」フォルダ内に置く
# config.py は1つ上の「カワウソマネージャー きなこver1.0」フォルダにある
# ──────────────────────────────────────────────────────────
import tkinter as tk
from tkinter import messagebox
import os
import re
import sys

# ── パス定義（セットアップフォルダの1つ上を参照）──────────
if getattr(sys, 'frozen', False):
    SETUP_DIR   = os.path.dirname(sys.executable)
    PROJECT_DIR = SETUP_DIR
else:
    SETUP_DIR   = os.path.dirname(os.path.abspath(__file__))
    PROJECT_DIR = os.path.normpath(os.path.join(SETUP_DIR, "..", ".."))
CONFIG_FILE = os.path.join(PROJECT_DIR, "config.py")

# Chrome 検索パス（LOCALAPPDATA 追加）
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        r"Google\Chrome\Application\chrome.exe"
    ),
]

# セキュリティ: 許可する URL プレフィックス（TikTok LiveCenter のみ）
ALLOWED_URL_PREFIX = "https://livecenter.tiktok.com/"

# ── 起動時チェック：config.py が存在するか確認 ────────────
if not os.path.exists(CONFIG_FILE):
    tk.Tk().withdraw()
    messagebox.showerror(
        "ファイルが見つかりません",
        f"config.py が見つかりません。\n\n"
        f"探した場所:\n{CONFIG_FILE}\n\n"
        f"「セットアップ」フォルダが正しい場所にあるか確認してください。"
    )
    sys.exit(1)

# セキュリティ: CONFIG_FILE が PROJECT_DIR 配下にあるか確認（パストラバーサル防止）
_real_config  = os.path.realpath(CONFIG_FILE)
_real_project = os.path.realpath(PROJECT_DIR)
if not _real_config.startswith(_real_project + os.sep):
    tk.Tk().withdraw()
    messagebox.showerror(
        "セキュリティエラー",
        "config.py のパスが不正です。\nセットアップを中止します。"
    )
    sys.exit(1)


# ── 入力バリデーション ────────────────────────────────────

def validate_tiktok_id(value: str) -> str | None:
    """正常なら None。エラーがあればメッセージ文字列を返す。"""
    if not value:
        return "TikTok ID を入力してください。"
    if len(value) > 24:
        return "TikTok ID は 24 文字以内で入力してください。"
    if not re.fullmatch(r"[a-zA-Z0-9_.]{1,24}", value):
        return (
            "TikTok ID に使えない文字が含まれています。\n"
            "（使用可能: 英数字・アンダースコア・ピリオドのみ）"
        )
    return None


def validate_url(value: str) -> str | None:
    """正常なら None。エラーがあればメッセージ文字列を返す。"""
    if not value:
        return "インサイトページ URL を入力してください。"
    if not value.startswith(ALLOWED_URL_PREFIX):
        return (
            f"URL は以下で始まる TikTok のアドレスのみ使用できます。\n\n"
            f"  {ALLOWED_URL_PREFIX}\n\n"
            f"デフォルト値:\n"
            f"  https://livecenter.tiktok.com/analytics/live_video?lang=ja-JP"
        )
    return None


# ── config.py から現在値を読み込む ────────────────────────

def read_config_value(key: str) -> str:
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                m = re.match(
                    rf'^{re.escape(key)}\s*=\s*["\'](.+?)["\']', line
                )
                if m:
                    return m.group(1)
    except Exception:
        pass
    return ""


# ── config.py の指定変数だけを書き換える ─────────────────

def update_config(tiktok_id: str, url: str) -> None:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        content = f.read()

    def replace_var(text: str, key: str, val: str) -> str:
        escaped_key = re.escape(key)
        pattern     = rf'^({escaped_key}\s*=\s*)["\'].*?["\']'
        safe_val    = repr(val)
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
    content = replace_var(content, "ANALYTICS_URL",       url)

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(content)


# ── Chrome のパスを探す ───────────────────────────────────

def find_chrome() -> str | None:
    for p in CHROME_PATHS:
        if p and os.path.isfile(p):
            return p
    return None


# ── 保存ボタン処理 ────────────────────────────────────────

def on_save() -> None:
    tiktok_id = var_id.get().strip().lstrip("@")
    url       = var_url.get().strip()

    # --- 入力バリデーション ---
    for err in (
        validate_tiktok_id(tiktok_id),
        validate_url(url),
    ):
        if err:
            messagebox.showwarning("入力エラー", err)
            return

    # --- Chrome チェック ---
    if not find_chrome():
        messagebox.showerror(
            "Chrome が見つかりません",
            "Google Chrome がインストールされているか確認してください。\n\n"
            "通常のインストール先:\n"
            r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        )
        return

    # --- 書き込み ---
    try:
        update_config(tiktok_id, url)
        messagebox.showinfo(
            "セットアップ完了 🎉",
            f"設定を保存しました！\n\n"
            f"  ・TikTok ID : @{tiktok_id}\n\n"
            f"保存先:\n{CONFIG_FILE}\n\n"
            f"──────────────────────────────\n"
            f"「カワウソマネージャー きなこ.bat」を\n"
            f"ダブルクリックして配信を開始してください！\n\n"
            f"※ 初回インサイト取得時に Chrome が起動します。\n"
            f"   TikTok にログインすると次回以降は自動ログインです。"
        )
    except PermissionError:
        messagebox.showerror(
            "書き込みエラー",
            f"config.py に書き込めませんでした。\n\n"
            f"対象ファイル:\n{CONFIG_FILE}\n\n"
            f"ファイルが他のアプリで開かれていないか確認してください。"
        )
    except Exception as e:
        messagebox.showerror("エラー", f"予期しないエラーが発生しました。\n\n{e}")


# ── GUI 構築 ──────────────────────────────────────────────

root = tk.Tk()
root.title("きなこ 初期セットアップ")
root.geometry("540x300")
root.configure(bg="#f0e6ff")
root.resizable(False, False)

# ヘッダー
tk.Label(
    root,
    text="🦦 カワウソマネージャー きなこ  初期セットアップ",
    font=("Meiryo", 12, "bold"),
    bg="#7c3aed", fg="white",
    padx=10, pady=12
).pack(fill="x")

# サブヘッダー
tk.Label(
    root,
    text="2項目を入力して「保存してセットアップ完了」を押してください",
    font=("Meiryo", 9),
    bg="#f0e6ff", fg="#555"
).pack(pady=(10, 2))

# config.py のパスを表示（確認用）
tk.Label(
    root,
    text=f"📄 config.py の場所: {CONFIG_FILE}",
    font=("Meiryo", 7),
    bg="#f0e6ff", fg="#999",
    wraplength=520
).pack(pady=(0, 6))

frame = tk.Frame(root, bg="#f0e6ff", padx=24, pady=6)
frame.pack(fill="both", expand=True)


def make_row(parent: tk.Frame, label: str, default: str = "") -> tk.StringVar:
    tk.Label(
        parent, text=label,
        bg="#f0e6ff", fg="#333",
        font=("Meiryo", 10, "bold"), anchor="w"
    ).pack(fill="x", pady=(8, 0))
    var = tk.StringVar(value=default)
    tk.Entry(
        parent, textvariable=var,
        font=("Meiryo", 10), width=58,
        relief="solid", bd=1
    ).pack(fill="x", pady=(2, 0))
    return var


var_id  = make_row(frame, "① TikTok ID（@ なし）",
                   read_config_value("MY_TIKTOK_USERNAME"))
var_url = make_row(frame, "② インサイトページ URL",
                   read_config_value("ANALYTICS_URL"))

tk.Button(
    frame,
    text="✅  保存してセットアップ完了",
    font=("Meiryo", 11, "bold"),
    bg="#7c3aed", fg="white",
    activebackground="#6d28d9", activeforeground="white",
    relief="flat", cursor="hand2",
    padx=10, pady=10,
    command=on_save
).pack(fill="x", pady=(16, 0))

# フッター
tk.Label(
    root,
    text="設定は config.py に保存　／　Chrome は自動検出・永続プロファイルで起動",
    font=("Meiryo", 7),
    bg="#f0e6ff", fg="#bbb"
).pack(pady=(6, 8))

root.mainloop()
