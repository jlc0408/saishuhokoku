"""
ブラウザ管理モジュール
Playwright を使って Edge を起動・操作する共通機能を提供する。
"""
from __future__ import annotations
import time
import config_manager

# タイムアウト設定（秒）
SAFE_WAIT = 0.5          # 操作間の安全待機
RETRY_INTERVAL = 60      # ボタン再探索間隔
MAX_RETRY = 4            # 最大再試行回数
STUCK_TIMEOUT = 300      # スタック判定時間（5分）


class BrowserError(Exception):
    """ブラウザ操作エラー（エラーコード付き）"""
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(f"[{code}] {message}")


class BrowserManager:
    def __init__(self, logger=None):
        self._playwright = None
        self._browser = None
        self._context = None
        self.page = None
        self.logger = logger

    def _log(self, msg: str):
        if self.logger:
            self.logger(msg)

    def launch(self):
        """Edgeを起動してメインページを返す。"""
        from playwright.sync_api import sync_playwright
        edge_path = config_manager.get("EDGE", "edge_path")
        self._log(f"Edge起動中: {edge_path}")
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(
            executable_path=edge_path,
            headless=False,
            args=["--disable-features=msSmartScreenProtection"],
        )
        self._context = self._browser.new_context()
        self.page = self._context.new_page()
        self._log("Edge起動完了")
        return self.page

    def close(self):
        try:
            if self._browser:
                self._browser.close()
            if self._playwright:
                self._playwright.stop()
        except Exception:
            pass

    def safe_wait(self):
        time.sleep(SAFE_WAIT)

    def wait_for_load(self, timeout: int = 30):
        """ページ読み込み完了まで待機。"""
        self.page.wait_for_load_state("domcontentloaded", timeout=timeout * 1000)
        self.safe_wait()

    def find_element_with_retry(self, selector: str, description: str = "要素"):
        """
        要素を探し、見つからない場合は1分おきに再探索する。
        4回連続失敗でBrowserErrorを送出する。
        """
        for attempt in range(MAX_RETRY):
            try:
                el = self.page.wait_for_selector(selector, timeout=RETRY_INTERVAL * 1000)
                return el
            except Exception:
                remaining = MAX_RETRY - attempt - 1
                if remaining > 0:
                    self._log(f"⚠ 「{description}」が見つかりません。1分後に再試行 (残り{remaining}回)")
                    time.sleep(RETRY_INTERVAL)
                else:
                    raise BrowserError(
                        "ERR-001",
                        f"「{description}」が{MAX_RETRY}回連続で見つかりませんでした。[selector: {selector}]",
                    )

    def click(self, selector: str, description: str = "ボタン"):
        """要素をクリックする（安全待機付き）。"""
        el = self.find_element_with_retry(selector, description)
        el.click()
        self.safe_wait()

    def fill(self, selector: str, value: str, description: str = "入力欄"):
        """入力欄に値をセットする（安全待機付き）。"""
        el = self.find_element_with_retry(selector, description)
        el.fill(value)
        self.safe_wait()

    def get_text(self, selector: str, description: str = "テキスト", default: str = "") -> str:
        """要素のテキストを取得する。取得失敗時はdefaultを返す。"""
        try:
            el = self.page.wait_for_selector(selector, timeout=10_000)
            return (el.inner_text() or "").strip()
        except Exception:
            self._log(f"⚠ 「{description}」の取得に失敗。空白で処理します。")
            return default

    def close_extra_tabs(self):
        """メインページ以外のタブを閉じる。"""
        pages = self._context.pages
        for p in pages:
            if p != self.page:
                p.close()
        self.safe_wait()

    def get_current_url(self) -> str:
        return self.page.url

    def navigate(self, url: str):
        self.page.goto(url)
        self.wait_for_load()
