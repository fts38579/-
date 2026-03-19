Copy
# ============================================================
#  インサイト手動取得.py  –  インサイトを今すぐ1回取得するだけ
# ============================================================
import os, sys, traceback, tkinter as tk
from tkinter import messagebox

# ── パス設定（EXE / .py 両対応）──────────────────────────────
if getattr(sys, 'frozen', False):
    _PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    _PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# ── ソース配下（非frozen時のみ）──────────────────────────────
if not getattr(sys, 'frozen', False):
    _SRC = os.path.join(_PROJECT_ROOT, "ソース")
    if _SRC not in sys.path:
        sys.path.insert(0, _SRC)


# ── メイン処理 ─────────────────────────────────────────────────
def main():
    root = tk.Tk()
    root.withdraw()

    result = messagebox.askokcancel(
        "インサイト手動取得",
        "TikTok LiveCenter のインサイトを今すぐ取得します。\n\n"
        "Chromeが自動的に起動し、最新の配信データを\n"
        "data/insights.xlsx に保存します。\n\n"
        "OK で開始、キャンセルで中断します。"
    )
    if not result:
        root.destroy()
        return

    # ── config 読み込み・バリデーション（専用エラー）─────────────
    try:
        import config
        config.validate()
    except Exception as e:
        messagebox.showerror(
            "❌ 設定エラー",
            f"config.py の設定に問題があります。\n\n"
            f"{e}\n\n"
            "・MY_TIKTOK_USERNAME が設定されているか確認してください。\n"
            "・ANALYTICS_URL が正しいか確認してください。"
        )
        root.destroy()
        return

    # ── インサイト取得 ───────────────────────────────────────────
    try:
        from modules.insights import collect_insights

        ok = collect_insights()

        if ok:
            messagebox.showinfo(
                "✅ 取得完了",
                "インサイトの取得が完了しました！\n"
                "保存先: data/insights.xlsx"
            )
        else:
            messagebox.showwarning(
                "⚠️ 取得失敗",
                "インサイトの取得に失敗しました。\n\n"
                "・TikTok にログインしているか確認してください。\n"
                "・data/debug_page.html で詳細を確認できます。"
            )

    except Exception as e:
        messagebox.showerror(
            "❌ エラー",
            f"インサイト取得中にエラーが発生しました。\n\n{e}\n\n"
            "・Chrome が既に起動していないか確認してください。\n"
            "・TikTok にログインしているか確認してください。"
        )
        traceback.print_exc()
    finally:
        root.destroy()


if __name__ == "__main__":
    main()

Copy