# ============================================================
#  modules/insights.py  修正版 v2
#  修正: _write_csv の拡張子変換ロジック整理、--disable-popup-blocking 削除
#  追加: webdriver-manager による ChromeDriver 自動管理 (Bug3)
# ============================================================
from __future__ import annotations

import os
import sys
import time
import datetime
import traceback
import csv as _csv

from bs4 import BeautifulSoup
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeDriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchWindowException,
)

# ── webdriver-manager（ChromeDriver 自動管理）────────────────────
try:
    from webdriver_manager.chrome import ChromeDriverManager
    _HAS_WDM = True
except ImportError:
    _HAS_WDM = False

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import config  # noqa: E402

# ────────────────────────────────────────────────────────────
WAIT_ROW_SEC     = 90
WAIT_DETAIL_SEC  = 60
WAIT_ELEMENT_SEC = 10

XPATH_FIRST_ROW  = '(//tr[contains(@class,"tt-live-table-row")])[1]'

COL_TITLE     = 0
COL_DATE      = 1
COL_DURATION  = 2
COL_VIEWERS   = 3
COL_FOLLOWERS = 4
COL_REWARD    = 5

DETAIL_KEYS = {
    "最高同時視聴者数": ["最高同時", "最高同時視聴", "Peak Viewers", "peak"],
    "平均視聴時間":     ["平均視聴時間", "平均視聴", "Avg Watch", "average watch"],
    "ギフト贈呈者数":   ["ギフト贈呈", "gift giver", "ギフター"],
    "LIVEおすすめ":    ["おすすめ", "recommend", "レコメンド"],
    "ダイヤ合計":      ["ダイヤ", "diamond", "diamonds"],
}


def _data_path(filename: str) -> str:
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, filename)


def _build_driver() -> ChromeDriver:
    options = ChromeOptions()
    options.add_argument(f"--user-data-dir={config.CHROME_PROFILE_DIR}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-notifications")
    # ★ --disable-popup-blocking を削除（意図しない新タブの抑制）
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # ── ChromeDriver 起動（webdriver-manager 優先）──────────────────
    if _HAS_WDM:
        try:
            print("[Insights] webdriver-manager で ChromeDriver を自動取得中…")
            service = ChromeService(ChromeDriverManager().install())
            driver  = ChromeDriver(service=service, options=options)
            print("[Insights] webdriver-manager: ChromeDriver 起動成功")
        except Exception as e_wdm:
            print(f"[Insights] webdriver-manager 失敗 ({e_wdm})、PATH の chromedriver を試みます")
            driver = ChromeDriver(options=options)
    else:
        print("[Insights] webdriver-manager 未インストール – PATH 上の chromedriver を使用")
        driver = ChromeDriver(options=options)

    driver.set_window_size(1366, 900)
    return driver


def _safe_text(el) -> str:
    try:
        return el.get_attribute("innerText").strip()
    except Exception:
        return ""


def _extract_list_row(driver: ChromeDriver) -> dict:
    result = {
        "LIVE名":        "N/A",
        "日付":          "N/A",
        "LIVE時間":      "N/A",
        "視聴数":        "N/A",
        "新規フォロワー": "N/A",
        "報酬":          "N/A",
    }
    try:
        wait      = WebDriverWait(driver, WAIT_ROW_SEC)
        first_row = wait.until(EC.presence_of_element_located(
            (By.XPATH, XPATH_FIRST_ROW)))
        tds = first_row.find_elements(By.TAG_NAME, "td")
        if len(tds) > COL_TITLE:     result["LIVE名"]         = _safe_text(tds[COL_TITLE])
        if len(tds) > COL_DATE:      result["日付"]           = _safe_text(tds[COL_DATE])
        if len(tds) > COL_DURATION:  result["LIVE時間"]       = _safe_text(tds[COL_DURATION])
        if len(tds) > COL_VIEWERS:   result["視聴数"]         = _safe_text(tds[COL_VIEWERS])
        if len(tds) > COL_FOLLOWERS: result["新規フォロワー"] = _safe_text(tds[COL_FOLLOWERS])
        if len(tds) > COL_REWARD:    result["報酬"]           = _safe_text(tds[COL_REWARD])
        print(f"[Insights] 一覧取得成功: {result['LIVE名']} / {result['日付']}")
    except TimeoutException:
        print("[Insights] 一覧行タイムアウト – N/A")
    except Exception as e:
        print(f"[Insights] 一覧取得エラー: {e}")
    return result


def _click_first_row(driver: ChromeDriver) -> bool:
    """
    最初の行をクリックして詳細ページへ遷移する。
    TikTokが新しいタブで開く場合はそちらへ切り替える。
    """
    try:
        wait = WebDriverWait(driver, WAIT_ROW_SEC)
        row  = wait.until(EC.element_to_be_clickable(
            (By.XPATH, XPATH_FIRST_ROW)))

        before_handles = set(driver.window_handles)
        current_url    = driver.current_url

        driver.execute_script("arguments[0].click();", row)
        time.sleep(2)

        after_handles = set(driver.window_handles)
        new_handles   = after_handles - before_handles
        if new_handles:
            driver.switch_to.window(list(new_handles)[0])
            print(f"[Insights] 新タブに切り替え: {driver.current_url}")
            return True

        try:
            WebDriverWait(driver, WAIT_DETAIL_SEC).until(
                lambda d: d.current_url != current_url)
            print(f"[Insights] 詳細ページ遷移成功: {driver.current_url}")
        except TimeoutException:
            print("[Insights] URL 変化なし – ページ内容で詳細を試みる")
        return True

    except NoSuchWindowException:
        print("[Insights] ウィンドウが閉じた – 利用可能なタブへ切り替え")
        try:
            if driver.window_handles:
                driver.switch_to.window(driver.window_handles[-1])
                print(f"[Insights] タブ切り替え成功: {driver.current_url}")
                return True
        except Exception as e2:
            print(f"[Insights] タブ切り替え失敗: {e2}")
        return False

    except TimeoutException:
        print("[Insights] 行クリック タイムアウト")
        return False
    except Exception as e:
        print(f"[Insights] 行クリック エラー: {e}")
        return False


def _extract_detail_metrics(driver: ChromeDriver) -> dict:
    result = {k: "N/A" for k in DETAIL_KEYS}
    time.sleep(3)
    try:
        soup       = BeautifulSoup(driver.page_source, "html.parser")
        text_nodes = soup.find_all(string=True)
        for metric_name, keywords in DETAIL_KEYS.items():
            for kw in keywords:
                kw_lower = kw.lower()
                for i, node in enumerate(text_nodes):
                    if kw_lower in node.lower():
                        for j in range(i + 1, min(i + 6, len(text_nodes))):
                            candidate = text_nodes[j].strip()
                            if candidate and any(c.isdigit() for c in candidate):
                                result[metric_name] = candidate
                                break
                    if result[metric_name] != "N/A":
                        break
                if result[metric_name] != "N/A":
                    break
        print(f"[Insights] 詳細取得: {result}")
    except Exception as e:
        print(f"[Insights] 詳細抽出エラー: {e}")
    return result


def _save_debug_html(driver: ChromeDriver, label: str = "") -> None:
    path = _data_path("data/debug_page.html")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"[Insights] デバッグ HTML 保存: {path} [{label}]")


def _write_csv(row_data: dict) -> None:
    """insights.csv へ追記保存する"""

    # ★ 拡張子変換ロジックを削除。config.CSV_INSIGHTS_FILE を直接参照
    csv_file = getattr(config, "CSV_INSIGHTS_FILE", "data/insights.csv")
    path     = _data_path(csv_file)
    dir_path = os.path.dirname(path)
    if dir_path:
        os.makedirs(dir_path, exist_ok=True)

    headers = [
        "取得日時", "LIVE名", "日付", "LIVE時間",
        "視聴数", "新規フォロワー", "報酬(ダイヤ)",
        "最高同時視聴者数", "平均視聴時間", "ギフト贈呈者数",
        "LIVEおすすめ", "ダイヤ合計",
    ]

    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = _csv.writer(f)
        if not file_exists:
            writer.writerow(headers)
        writer.writerow([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            row_data.get("LIVE名",           "N/A"),
            row_data.get("日付",             "N/A"),
            row_data.get("LIVE時間",         "N/A"),
            row_data.get("視聴数",           "N/A"),
            row_data.get("新規フォロワー",   "N/A"),
            row_data.get("報酬",             "N/A"),
            row_data.get("最高同時視聴者数", "N/A"),
            row_data.get("平均視聴時間",     "N/A"),
            row_data.get("ギフト贈呈者数",   "N/A"),
            row_data.get("LIVEおすすめ",     "N/A"),
            row_data.get("ダイヤ合計",       "N/A"),
        ])
    print(f"[Insights] CSV 保存完了: {path}")


def collect_insights() -> bool:
    print("[Insights] ====== インサイト取得 開始 ======")
    driver: ChromeDriver | None = None
    success = False

    try:
        print("[Insights] Chrome 起動中（ポート開放なし）…")
        driver = _build_driver()

        print(f"[Insights] {config.ANALYTICS_URL} へ移動中…")
        driver.get(config.ANALYTICS_URL)

        list_data = _extract_list_row(driver)
        _save_debug_html(driver, "list-page")

        click_ok    = _click_first_row(driver)
        detail_data = {k: "N/A" for k in DETAIL_KEYS}

        if click_ok:
            _save_debug_html(driver, "detail-page")
            detail_data = _extract_detail_metrics(driver)
        else:
            print("[Insights] クリック失敗 – 詳細データは N/A で保存")

        _write_csv({**list_data, **detail_data})
        success = True

    except WebDriverException as e:
        print(f"[Insights] WebDriver エラー: {e}")
        traceback.print_exc()
    except Exception as e:
        print(f"[Insights] 予期しないエラー: {e}")
        traceback.print_exc()
    finally:
        if driver:
            try:
                driver.quit()
                print("[Insights] Chrome 正常終了")
            except Exception:
                pass
        print("[Insights] ====== インサイト取得 完了 ======")

    return success
