# きなこのレポート.py  完全版（タブUI + エクスポートボタン右上配置）
import os, sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
import io
import re

# ─── パス設定（PyInstaller対応）──────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DATA_DIR     = os.path.join(BASE_DIR, "data")
CSV_FILE     = os.path.join(DATA_DIR, "gift_timeline.csv")
VIEWERS_FILE = os.path.join(DATA_DIR, "viewers.csv")

try:
    _repo_root = BASE_DIR
    if _repo_root not in sys.path:
        sys.path.insert(0, _repo_root)
    import config as _cfg
    _raw = getattr(_cfg, "CSV_INSIGHTS_FILE", "data/insights.csv")
except Exception:
    _raw = "data/insights.csv"
INSIGHTS_CSV = os.path.join(DATA_DIR, os.path.splitext(os.path.basename(_raw))[0] + ".csv")

# ─── tkcalendar インポート ────────────────────────────────
try:
    from tkcalendar import DateEntry
    _HAS_CALENDAR = True
except ImportError:
    _HAS_CALENDAR = False

# ─── 日本語フォント設定 ──────────────────────────────────
def set_japanese_font():
    candidates = ["Meiryo", "MS Gothic", "Yu Gothic", "IPAGothic", "Noto Sans CJK JP"]
    for name in candidates:
        for f in fm.fontManager.ttflist:
            if name.lower() in f.name.lower():
                matplotlib.rcParams["font.family"] = f.name
                return
    for f in fm.fontManager.ttflist:
        if any(c in f.name for c in ["Gothic", "Mincho", "Hiragino", "Noto"]):
            matplotlib.rcParams["font.family"] = f.name
            return

set_japanese_font()
matplotlib.rcParams["axes.unicode_minus"] = False

# ─── データ読み込み ──────────────────────────────────────
def load_insights():
    if not os.path.exists(INSIGHTS_CSV):
        return None, f"insights.csv が見つかりません。\nパス: {INSIGHTS_CSV}"
    try:
        df = pd.read_csv(INSIGHTS_CSV, encoding="utf-8-sig")
        date_col = None
        if "取得日時" in df.columns:
            date_col = "取得日時"
        else:
            for col in df.columns:
                if "日" in col or "date" in col.lower() or "時" in col or "取得" in col:
                    date_col = col
                    break
        if date_col:
            df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
        else:
            try:
                df["_date"] = pd.to_datetime(df.iloc[:, 0], errors="coerce")
            except Exception:
                return None, "日付列を認識できませんでした。"
        df = df.dropna(subset=["_date"])
        df = df.sort_values("_date").reset_index(drop=True)
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

        parsed               = df_gift["detail"].apply(parse_detail)
        df_gift              = df_gift.copy()
        df_gift["gift_name"] = [p[0] for p in parsed]
        df_gift["count"]     = [p[1] for p in parsed]
        df_gift["_date"]     = df_gift["timestamp"]
        return df_gift, None
    except Exception as e:
        return None, str(e)

def load_viewers():
    if not os.path.exists(VIEWERS_FILE):
        return None, (
            "viewers.csv が見つかりません。\n"
            "配信を1回以上終了すると自動生成されます。"
        )
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

def find_col(df, *keywords):
    for kw in keywords:
        for col in df.columns:
            if kw in col:
                return col
    return None

# ─── 日付入力ウィジェット生成 ────────────────────────────
def _make_date_entry(parent, var: tk.StringVar):
    if _HAS_CALENDAR:
        today = datetime.today()
        try:
            init_date = datetime.strptime(var.get(), "%Y-%m-%d")
        except Exception:
            init_date = today
        entry = DateEntry(
            parent,
            textvariable=var,
            font=("Meiryo", 11),
            width=14,
            date_pattern="yyyy-mm-dd",
            year=init_date.year,
            month=init_date.month,
            day=init_date.day,
            background="#7c3aed",
            foreground="white",
            borderwidth=1,
        )
    else:
        entry = tk.Entry(parent, textvariable=var,
                         font=("Meiryo", 11), width=16,
                         relief="solid", bd=1)
    return entry

# ─── グラフをタブフレーム内に埋め込む ───────────────────
def _embed_chart(frame, fig):
    for widget in frame.winfo_children():
        widget.destroy()
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)
    return canvas

# ─── Excelエクスポート ───────────────────────────────────
def export_excel(fig, df, default_title):
    if fig is None or df is None:
        messagebox.showwarning("未表示", "先にグラフを表示してからエクスポートしてください。")
        return
    try:
        import openpyxl
        from openpyxl.drawing.image import Image as XLImage
    except ImportError:
        messagebox.showerror(
            "ライブラリエラー",
            "openpyxl がインストールされていません。\n"
            "  py -m pip install openpyxl"
        )
        return
    safe_title = default_title.replace("（", "_").replace("）", "").replace(" ", "").replace("～", "-")
    path = filedialog.asksaveasfilename(
        title="Excelファイルを保存",
        initialfile=f"{safe_title}.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
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
            ws_data.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)
        ws_chart = wb.create_sheet(title="グラフ")
        img_buf  = io.BytesIO()
        fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
        img_buf.seek(0)
        ws_chart.add_image(XLImage(img_buf), "A1")
        wb.save(path)
        messagebox.showinfo("保存完了", f"Excelファイルを保存しました。\n{path}")
    except Exception as e:
        messagebox.showerror("保存エラー", f"Excel保存に失敗しました。\n{e}")

# ─── CSVエクスポート ─────────────────────────────────────
def export_csv(df, default_title):
    if df is None:
        messagebox.showwarning("未表示", "先にグラフを表示してからエクスポートしてください。")
        return
    default_name = (
        default_title.replace("（", "_").replace("）", "")
                     .replace(" ", "").replace("～", "-") + ".csv"
    )
    path = filedialog.asksaveasfilename(
        title="CSVを保存",
        initialfile=default_name,
        defaultextension=".csv",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
    )
    if not path:
        return
    try:
        out = df.copy()
        for col in out.columns:
            if pd.api.types.is_datetime64_any_dtype(out[col]):
                out[col] = out[col].astype(str)
        out.to_csv(path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("保存完了", f"CSVを保存しました。\n{path}")
    except Exception as e:
        messagebox.showerror("保存エラー", f"保存に失敗しました。\n{e}")

# ─── 現在タブのfig/df/titleを取得 ───────────────────────
def _get_current_fig_df():
    idx = notebook.index(notebook.select())
    if idx == 0:
        title = f"インサイト（{var_ins_start.get()} ～ {var_ins_end.get()}）"
        return _insight_fig, _insight_df, title
    elif idx == 1:
        title = f"ギフトタイムライン（{var_gift_start.get()} ～ {var_gift_end.get()}）"
        return _gift_fig, _gift_df, title
    else:
        return _repeat_fig, _repeat_df, "リピート率レポート"

def on_export_excel():
    fig, df, title = _get_current_fig_df()
    export_excel(fig, df, title)

def on_export_csv():
    fig, df, title = _get_current_fig_df()
    export_csv(df, title)

# ════════════════════════════════════════════════════════════
#  メイン GUI（先に root を作る）
# ════════════════════════════════════════════════════════════
root = tk.Tk()
root.title("きなこのレポート")
root.geometry("900x750")
root.configure(bg="#f0e6ff")
root.resizable(True, True)

# ── ヘッダー（タイトル左・エクスポートボタン右）──────────
header_frame = tk.Frame(root, bg="#7c3aed")
header_frame.pack(fill="x")

tk.Label(header_frame,
         text="🦦 カワウソマネージャー きなこ",
         font=("Meiryo", 14, "bold"),
         bg="#7c3aed", fg="white",
         padx=14, pady=10).pack(side="left")

# ★ エクスポートボタンを右上に配置
btn_export_frame = tk.Frame(header_frame, bg="#7c3aed")
btn_export_frame.pack(side="right", padx=14, pady=6)

tk.Button(btn_export_frame,
          text="📊 Excelエクスポート",
          command=on_export_excel,
          font=("Meiryo", 9, "bold"),
          bg="#1d4ed8", fg="white", relief="flat",
          activebackground="#1e3a8a", activeforeground="white",
          padx=12, pady=5).pack(side="left", padx=(0, 8))

tk.Button(btn_export_frame,
          text="📥 CSVエクスポート",
          command=on_export_csv,
          font=("Meiryo", 9, "bold"),
          bg="#059669", fg="white", relief="flat",
          activebackground="#047857", activeforeground="white",
          padx=12, pady=5).pack(side="left")

# ── タブスタイル設定 ──────────────────────────────────────
style = ttk.Style()
style.theme_use("default")
style.configure("TNotebook",
                background="#f0e6ff",
                borderwidth=0)
style.configure("TNotebook.Tab",
                font=("Meiryo", 11, "bold"),
                padding=[20, 8],
                background="#d8b4fe",
                foreground="#4c1d95")
style.map("TNotebook.Tab",
          background=[("selected", "#7c3aed")],
          foreground=[("selected", "white")])
style.configure("TFrame", background="#f0e6ff")

# ── タブ本体 ──────────────────────────────────────────────
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))

tab_insight = ttk.Frame(notebook)
tab_gift    = ttk.Frame(notebook)
tab_repeat  = ttk.Frame(notebook)

notebook.add(tab_insight, text="  📊 インサイト  ")
notebook.add(tab_gift,    text="  🎁 ギフト  ")
notebook.add(tab_repeat,  text="  👥 リピート率  ")

# ════════════════════════════════════════════════════════════
#  タブ① インサイト
# ════════════════════════════════════════════════════════════
_insight_fig = None
_insight_df  = None

def _build_insights_fig(df, fig, axes, title):
    col_peak    = "最高同時視聴者数" if "最高同時視聴者数" in df.columns else \
                  find_col(df, "最高同時", "peak", "同接")
    col_diamond = "ダイヤ合計"      if "ダイヤ合計"      in df.columns else \
                  find_col(df, "diamond")
    col_gifter  = "ギフト贈呈者数"  if "ギフト贈呈者数"  in df.columns else \
                  find_col(df, "ギフト贈呈", "gifter")
    col_watch   = "平均視聴時間"    if "平均視聴時間"    in df.columns else \
                  find_col(df, "平均視聴", "watch", "view")

    for col in [col_peak, col_diamond, col_gifter, col_watch]:
        if col:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    diamond_total = int(df[col_diamond].sum(skipna=True)) if col_diamond else 0
    fig.suptitle(title + f"  ◆ 期間合計ダイヤ: {diamond_total:,}", fontsize=11, y=0.98)

    plot_configs = [
        (axes[0][0], col_peak,    "#4f86c6", "最高同時視聴者数（人）"),
        (axes[0][1], col_diamond, "#f5a623", "ダイヤ数"),
        (axes[1][0], col_gifter,  "#7ed321", "ギフト贈呈者数（人）"),
        (axes[1][1], col_watch,   "#e87c7c", "平均視聴時間"),
    ]
    for ax, col, color, ylabel in plot_configs:
        if col and col in df.columns:
            mask        = df[col].notna()
            vals        = df.loc[mask, col]
            date_labels = (df.loc[mask, "_date"].dt.strftime("%m/%d")
                           if "_date" in df.columns
                           else [str(i) for i in range(len(vals))])
            if vals.empty:
                ax.text(0.5, 0.5, "データなし", ha="center", va="center",
                        transform=ax.transAxes, fontsize=11)
            else:
                ax.bar(range(len(vals)), vals, color=color, alpha=0.8)
                ax.set_xticks(range(len(vals)))
                ax.set_xticklabels(date_labels, rotation=45, fontsize=8)
                ax.set_ylabel(ylabel, fontsize=9)
                mean_val = vals.mean()
                ax.axhline(mean_val, color="red", linestyle="--", linewidth=1.2,
                           label=f"平均: {mean_val:.1f}")
                ax.legend(fontsize=8)
        else:
            ax.text(0.5, 0.5, "データなし", ha="center", va="center",
                    transform=ax.transAxes, fontsize=11)
        ax.set_title(ylabel, fontsize=10)
    plt.tight_layout(rect=[0, 0, 1, 0.94])

def on_show_insights():
    global _insight_fig, _insight_df
    df, err = load_insights()
    if err:
        messagebox.showerror("エラー", err)
        return
    try:
        start = pd.to_datetime(var_ins_start.get())
        end   = pd.to_datetime(var_ins_end.get()) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    except Exception:
        messagebox.showerror("入力エラー", "日付形式が正しくありません")
        return
    df = df[(df["_date"] >= start) & (df["_date"] <= end)]
    if df.empty:
        messagebox.showinfo("データなし", "指定期間にデータがありません。")
        return
    title = f"インサイト（{var_ins_start.get()} ～ {var_ins_end.get()}）"
    plt.close("all")
    fig, axes = plt.subplots(2, 2, figsize=(10, 6))
    _build_insights_fig(df, fig, axes, title)
    _insight_fig = fig
    _insight_df  = df
    _embed_chart(frame_ins_graph, fig)

def _build_insight_tab(parent):
    global var_ins_start, var_ins_end, frame_ins_graph

    today     = datetime.today()
    one_month = today - timedelta(days=30)

    ctrl = tk.Frame(parent, bg="#f0e6ff")
    ctrl.pack(fill="x", padx=20, pady=(12, 6))

    tk.Label(ctrl, text="開始日", bg="#f0e6ff",
             font=("Meiryo", 10)).grid(row=0, column=0, padx=(0, 6))
    var_ins_start = tk.StringVar(value=one_month.strftime("%Y-%m-%d"))
    _make_date_entry(ctrl, var_ins_start).grid(row=0, column=1, padx=(0, 20))

    tk.Label(ctrl, text="終了日", bg="#f0e6ff",
             font=("Meiryo", 10)).grid(row=0, column=2, padx=(0, 6))
    var_ins_end = tk.StringVar(value=today.strftime("%Y-%m-%d"))
    _make_date_entry(ctrl, var_ins_end).grid(row=0, column=3, padx=(0, 20))

    tk.Button(ctrl, text="グラフを表示",
              command=on_show_insights,
              font=("Meiryo", 10, "bold"),
              bg="#7c3aed", fg="white", relief="flat",
              activebackground="#5b21b6", activeforeground="white",
              padx=16, pady=4).grid(row=0, column=4)

    frame_ins_graph = tk.Frame(parent, bg="#f0e6ff")
    frame_ins_graph.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    parent.after(100, on_show_insights)

# ════════════════════════════════════════════════════════════
#  タブ② ギフトタイムライン
# ════════════════════════════════════════════════════════════
_gift_fig = None
_gift_df  = None

def _build_gift_fig(df_range, period_str, fig, axes):
    unique_gifters = df_range["user"].nunique() if "user" in df_range.columns else 0
    gift_count     = len(df_range)
    summary_str    = f"  ◆ ギフター: {unique_gifters}人  |  ギフト回数: {gift_count}回"
    fig.suptitle(f"ギフトタイムライン（{period_str}）" + summary_str, fontsize=11, y=0.98)

    ax1 = axes[0]
    if "_date" in df_range.columns and not df_range.empty:
        df_range = df_range.copy()
        df_range["hour"] = df_range["_date"].dt.hour
        hourly = df_range.groupby("hour").size()
        ax1.bar(hourly.index, hourly.values, color="#f5a623", alpha=0.8)
        ax1.set_xlabel("時刻（時）", fontsize=9)
        ax1.set_ylabel("ギフト回数", fontsize=9)
    else:
        ax1.text(0.5, 0.5, "データなし", ha="center", va="center",
                 transform=ax1.transAxes)
    ax1.set_title("時間帯別ギフト回数", fontsize=10)

    ax2 = axes[1]
    if "user" in df_range.columns and not df_range.empty:
        top_gifters = df_range.groupby("user").size().nlargest(10)
        ax2.barh(top_gifters.index[::-1], top_gifters.values[::-1],
                 color="#7ed321", alpha=0.8)
        ax2.set_xlabel("ギフト回数", fontsize=9)
    else:
        ax2.text(0.5, 0.5, "データなし", ha="center", va="center",
                 transform=ax2.transAxes)
    ax2.set_title("トップギフター Top10", fontsize=10)

    ax3 = axes[2]
    if "gift_name" in df_range.columns and not df_range.empty:
        top_gifts = df_range.groupby("gift_name")["count"].sum().nlargest(10)
        ax3.barh(top_gifts.index[::-1], top_gifts.values[::-1],
                 color="#4f86c6", alpha=0.8)
        ax3.set_xlabel("合計個数", fontsize=9)
    else:
        ax3.text(0.5, 0.5, "データなし", ha="center", va="center",
                 transform=ax3.transAxes)
    ax3.set_title("ギフト種別 Top10", fontsize=10)
    plt.tight_layout(rect=[0, 0, 1, 0.94])

def on_show_gift():
    global _gift_fig, _gift_df
    df, err = load_gifts()
    if err:
        messagebox.showerror("エラー", err)
        return
    if df is None or df.empty:
        messagebox.showinfo("データなし", "gift_timeline.csv にギフトデータがありません。")
        return
    try:
        start = pd.to_datetime(var_gift_start.get()).date()
        end   = (pd.to_datetime(var_gift_end.get()) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)).date()
    except Exception:
        messagebox.showerror("入力エラー", "日付形式が正しくありません")
        return
    df_range = df[(df["_date"].dt.date >= start) & (df["_date"].dt.date <= end)].copy()
    if df_range.empty:
        messagebox.showinfo("データなし", f"{var_gift_start.get()} ～ {var_gift_end.get()} のギフトデータがありません。")
        return
    period_str = f"{var_gift_start.get()} ～ {var_gift_end.get()}"
    plt.close("all")
    fig, axes = plt.subplots(1, 3, figsize=(12, 5))
    _build_gift_fig(df_range, period_str, fig, axes)
    _gift_fig = fig
    _gift_df  = df_range
    _embed_chart(frame_gift_graph, fig)

def _build_gift_tab(parent):
    global var_gift_start, var_gift_end, frame_gift_graph

    today     = datetime.today()
    one_month = today - timedelta(days=30)

    ctrl = tk.Frame(parent, bg="#f0e6ff")
    ctrl.pack(fill="x", padx=20, pady=(12, 6))

    tk.Label(ctrl, text="開始日", bg="#f0e6ff",
             font=("Meiryo", 10)).grid(row=0, column=0, padx=(0, 6))
    var_gift_start = tk.StringVar(value=one_month.strftime("%Y-%m-%d"))
    _make_date_entry(ctrl, var_gift_start).grid(row=0, column=1, padx=(0, 20))

    tk.Label(ctrl, text="終了日", bg="#f0e6ff",
             font=("Meiryo", 10)).grid(row=0, column=2, padx=(0, 6))
    var_gift_end = tk.StringVar(value=today.strftime("%Y-%m-%d"))
    _make_date_entry(ctrl, var_gift_end).grid(row=0, column=3, padx=(0, 20))

    tk.Button(ctrl, text="グラフを表示",
              command=on_show_gift,
              font=("Meiryo", 10, "bold"),
              bg="#7c3aed", fg="white", relief="flat",
              activebackground="#5b21b6", activeforeground="white",
              padx=16, pady=4).grid(row=0, column=4)

    frame_gift_graph = tk.Frame(parent, bg="#f0e6ff")
    frame_gift_graph.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    parent.after(150, on_show_gift)

# ════════════════════════════════════════════════════════════
#  タブ③ リピート率
# ════════════════════════════════════════════════════════════
_repeat_fig = None
_repeat_df  = None

def _build_repeat_fig(df, fig, axes):
    uid_col  = "uid" if "uid" in df.columns else df.columns[2] if len(df.columns) > 2 else None
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

    title_str = (f"リピート率レポート  |  ユニーク視聴者: {total_unique}人  "
                 f"リピーター: {repeaters}人  リピート率: {repeat_rate:.1f}%")
    fig.suptitle(title_str, fontsize=11, y=0.98)

    ax0 = axes[0]
    if total_unique > 0:
        ax0.pie([repeaters, first_only],
                labels=[f"リピーター\n{repeaters}人", f"初回のみ\n{first_only}人"],
                colors=["#7c3aed", "#c4b5fd"],
                autopct="%1.1f%%", startangle=90,
                textprops={"fontsize": 10})
    else:
        ax0.text(0.5, 0.5, "データなし", ha="center", va="center",
                 transform=ax0.transAxes, fontsize=11)
    ax0.set_title("リピーター比率", fontsize=11)

    ax1 = axes[1]
    if not session_viewers.empty:
        dates = [str(d) for d in session_viewers.index]
        ax1.bar(range(len(dates)), session_viewers.values, color="#4f86c6", alpha=0.8)
        ax1.set_xticks(range(len(dates)))
        ax1.set_xticklabels(dates, rotation=45, fontsize=8)
        ax1.set_ylabel("ユニーク視聴者数（人）", fontsize=9)
        mean_v = session_viewers.mean()
        ax1.axhline(mean_v, color="red", linestyle="--", linewidth=1.2,
                    label=f"平均: {mean_v:.1f}")
        ax1.legend(fontsize=8)
    else:
        ax1.text(0.5, 0.5, "データなし", ha="center", va="center",
                 transform=ax1.transAxes, fontsize=11)
    ax1.set_title("セッション別ユニーク視聴者", fontsize=10)

    ax2 = axes[2]
    if len(top_repeaters) > 0:
        ax2.barh(top_labels[::-1], top_repeaters.values[::-1],
                 color="#7ed321", alpha=0.8)
        ax2.set_xlabel("参加セッション数", fontsize=9)
    else:
        ax2.text(0.5, 0.5, "リピーターなし", ha="center", va="center",
                 transform=ax2.transAxes, fontsize=11)
    ax2.set_title("リピーター Top10（参加セッション数）", fontsize=10)
    plt.tight_layout(rect=[0, 0, 1, 0.94])

def on_show_repeat():
    global _repeat_fig, _repeat_df
    df, err = load_viewers()
    if err:
        messagebox.showerror("エラー", err)
        return
    if df.empty:
        messagebox.showinfo("データなし", "viewers.csv にデータがありません。")
        return
    try:
        start = pd.to_datetime(var_rep_start.get()).date()
        end   = (pd.to_datetime(var_rep_end.get()) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)).date()
        df    = df[(df["session_date"] >= start) & (df["session_date"] <= end)]
    except Exception:
        messagebox.showerror("入力エラー", "日付形式が正しくありません")
        return
    if df.empty:
        messagebox.showinfo("データなし", "指定期間にデータがありません。")
        return
    plt.close("all")
    fig, axes = plt.subplots(1, 3, figsize=(12, 5))
    _build_repeat_fig(df, fig, axes)
    _repeat_fig = fig
    _repeat_df  = df
    _embed_chart(frame_rep_graph, fig)

def _build_repeat_tab(parent):
    global var_rep_start, var_rep_end, frame_rep_graph

    today     = datetime.today()
    one_month = today - timedelta(days=30)

    ctrl = tk.Frame(parent, bg="#f0e6ff")
    ctrl.pack(fill="x", padx=20, pady=(12, 6))

    tk.Label(ctrl, text="開始日", bg="#f0e6ff",
             font=("Meiryo", 10)).grid(row=0, column=0, padx=(0, 6))
    var_rep_start = tk.StringVar(value=one_month.strftime("%Y-%m-%d"))
    _make_date_entry(ctrl, var_rep_start).grid(row=0, column=1, padx=(0, 20))

    tk.Label(ctrl, text="終了日", bg="#f0e6ff",
             font=("Meiryo", 10)).grid(row=0, column=2, padx=(0, 6))
    var_rep_end = tk.StringVar(value=today.strftime("%Y-%m-%d"))
    _make_date_entry(ctrl, var_rep_end).grid(row=0, column=3, padx=(0, 20))

    tk.Button(ctrl, text="グラフを表示",
              command=on_show_repeat,
              font=("Meiryo", 10, "bold"),
              bg="#7c3aed", fg="white", relief="flat",
              activebackground="#5b21b6", activeforeground="white",
              padx=16, pady=4).grid(row=0, column=4)

    frame_rep_graph = tk.Frame(parent, bg="#f0e6ff")
    frame_rep_graph.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    parent.after(200, on_show_repeat)

# ── タブ内容を構築 ────────────────────────────────────────
_build_insight_tab(tab_insight)
_build_gift_tab(tab_gift)
_build_repeat_tab(tab_repeat)

# ── フッター ──────────────────────────────────────────────
tk.Label(root, text=f"data フォルダ: {DATA_DIR}",
         font=("Meiryo", 8),
         bg="#f0e6ff", fg="#999").pack(side="bottom", pady=6)

root.mainloop()
