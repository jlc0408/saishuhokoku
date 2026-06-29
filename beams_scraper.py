"""
BEAMS操作モジュール
ログイン・検索・案件データ取得を担当する。
"""
from __future__ import annotations
import re
import time
import config_manager
from browser_manager import BrowserManager, BrowserError

BEAMS_URL = "https://ap.salesforce.com/secur/login_portal.jsp?orgId=00D10000000IDQS&portalId=06010000000Lc5O"

SEL_USERNAME     = "#username"
SEL_PASSWORD     = "#password"
SEL_LOGIN_BTN    = "input.btn[value='Login']"
SEL_SEARCH_TAB   = "a[title='申込検索Bタブ']"
SEL_RECORD_COUNT = ".recordCounter"

# colons を含むIDはgetElementByIdで操作する
ID_DATE_FROM  = "j_id0:sve_form1:pageBlock1:searchPageBlockSection1:inputSearchFromTo5_from"
ID_DATE_TO    = "j_id0:sve_form1:pageBlock1:searchPageBlockSection1:inputSearchFromTo5_to"
ID_SEARCH_BTN = "j_id0:sve_form1:pageBlock1:j_id75:searchButton1"


def _date_to_str(date_obj) -> str:
    return date_obj.strftime("%Y/%m/%d")


def _js_get_next_td_text(page, label_text: str) -> str:
    """
    td.labelCol のテキストが label_text に一致する行の
    次の兄弟 td のテキストを JS で直接取得して返す。
    見つからない場合は空文字。
    """
    script = f"""
    (() => {{
        const labels = document.querySelectorAll('td.labelCol');
        for (const td of labels) {{
            if (td.innerText && td.innerText.includes('{label_text}')) {{
                const next = td.nextElementSibling;
                return next ? next.innerText.trim() : '';
            }}
        }}
        return '';
    }})()
    """
    try:
        result = page.evaluate(script)
        return (result or "").strip()
    except Exception:
        return ""


class BeamsScraper:
    def __init__(self, bm: BrowserManager, logger=None):
        self.bm = bm
        self.logger = logger

    def _log(self, msg: str):
        if self.logger:
            self.logger(msg)

    # ──────────────────────────────────────────
    # ⑤ ログイン
    # ──────────────────────────────────────────
    def login(self):
        url = config_manager.get("BEAMS", "url") or BEAMS_URL
        username = config_manager.get("BEAMS", "username")
        password = config_manager.get("BEAMS", "password")

        self._log(f"BEAMSログインページへ移動: {url}")
        self.bm.navigate(url)

        self._log("ユーザー名・パスワードを入力")
        self.bm.fill(SEL_USERNAME, username, "ユーザー名入力欄")
        self.bm.fill(SEL_PASSWORD, password, "パスワード入力欄")
        self.bm.click(SEL_LOGIN_BTN, "ログインボタン")
        self.bm.wait_for_load()

        self.bm.close_extra_tabs()
        current = self.bm.get_current_url()
        self._log(f"ログイン後URL: {current}")

        # ログイン失敗メッセージを検出
        try:
            error_el = self.bm.page.query_selector("div.errorMsg")
            if error_el:
                error_text = (error_el.inner_text() or "").strip()
                if error_text:
                    raise BrowserError(
                        "ERR-LOGIN",
                        "ログインに失敗しました。IDまたはパスワードが正しくありません。\n"
                        "「設定」タブでIDとパスワードを確認・更新してください。"
                    )
        except BrowserError:
            raise
        except Exception:
            pass

        if "login_portal" in current or "/login" in current.lower():
            self._log("⚠ ログイン後URLが不正。ホームへ再移動")
            self.bm.navigate("https://ap.salesforce.com/home/home.jsp")

        self._log("ログイン完了")

    # ──────────────────────────────────────────
    # ⑥ 申込検索Bタブへ移動
    # ──────────────────────────────────────────
    def go_to_search_tab(self):
        self._log("申込検索Bタブへ移動")
        self.bm.click(SEL_SEARCH_TAB, "申込検索Bタブ")
        self.bm.wait_for_load()

    # ──────────────────────────────────────────
    # ⑦⑧ 日付検索
    # ──────────────────────────────────────────
    def search_by_date(self, date_obj) -> int:
        date_str = _date_to_str(date_obj)
        self._log(f"検索日付: {date_str}")

        # colons含むIDのためgetElementByIdで値をセット
        self.bm.page.evaluate(
            f"document.getElementById('{ID_DATE_FROM}').value = '{date_str}'"
        )
        self.bm.safe_wait()
        self.bm.page.evaluate(
            f"document.getElementById('{ID_DATE_TO}').value = '{date_str}'"
        )
        self.bm.safe_wait()

        self._log("検索ボタンをクリック")
        self.bm.page.evaluate(f"document.getElementById('{ID_SEARCH_BTN}').click()")
        self.bm.wait_for_load()

        # ⑨ 件数取得
        count_text = self.bm.get_text(SEL_RECORD_COUNT, "件数")
        count = 0
        m = re.search(r"\d+", count_text)
        if m:
            count = int(m.group())
        self._log(f"検索結果: {count}件")
        return count

    # ──────────────────────────────────────────
    # ⑩ 案件リスト取得
    # ──────────────────────────────────────────
    def _collect_req_on_page(self, seen: set) -> list[dict]:
        """現在ページに表示されているREQリンクを収集して返す（重複除外）。"""
        links = self.bm.page.query_selector_all("a[href*='salesforce.com/a2C']")
        results = []
        for link in links:
            text = (link.inner_text() or "").strip()
            href = link.get_attribute("href") or ""
            if text.startswith("REQ") and text not in seen:
                seen.add(text)
                if href.startswith("/"):
                    href = "https://ap.salesforce.com" + href
                results.append({"req": text, "url": href})
        return results

    def _get_page_info(self) -> tuple[int, int]:
        """
        ページャーの pagedisplay から「現在ページ / 総ページ数」を返す。
        取得できない場合は (1, 1) を返す。
        """
        try:
            val = self.bm.page.eval_on_selector(
                "input.pagedisplay", "el => el.value"
            ) or ""
            # "1/6" 形式
            m = re.match(r"(\d+)\s*/\s*(\d+)", val.strip())
            if m:
                return int(m.group(1)), int(m.group(2))
        except Exception:
            pass
        return 1, 1

    def _go_next_page(self):
        """次ページボタンをクリックして読み込み完了まで待機する。"""
        self.bm.page.click("img.next")
        self.bm.wait_for_load()

    def get_req_list(self) -> list[dict]:
        """全ページを巡回してREQリストを取得する（重複除外）。"""
        self._log("案件リスト取得中")
        results = []
        seen: set[str] = set()

        current_page, total_pages = self._get_page_info()
        self._log(f"  ページ {current_page}/{total_pages}")
        results.extend(self._collect_req_on_page(seen))

        while current_page < total_pages:
            self._log(f"  次のページへ移動 ({current_page + 1}/{total_pages})")
            self._go_next_page()
            current_page, total_pages = self._get_page_info()
            self._log(f"  ページ {current_page}/{total_pages}")
            results.extend(self._collect_req_on_page(seen))

        self._log(f"{len(results)}件のREQを取得（全{total_pages}ページ）")
        return results

    # ──────────────────────────────────────────
    # ⑬〜⑱ 案件詳細ページからデータ取得
    # ──────────────────────────────────────────
    def get_detail(self, req_url: str) -> dict:
        self._log(f"詳細ページを別タブで開く: {req_url}")
        context = self.bm._context
        detail_page = context.new_page()
        detail_page.goto(req_url)
        detail_page.wait_for_load_state("domcontentloaded")
        time.sleep(0.5)

        data = {}

        # ── 三次店（代理店様用フリーボックス①）──
        # JS内でラベルspanを探し、同一tr内のtd.dataColのテキストを返す
        mitsugiten_raw = ""
        try:
            mitsugiten_raw = detail_page.evaluate("""
            (() => {
                const spans = document.querySelectorAll('span.helpButton');
                for (const span of spans) {
                    if (span.innerText && span.innerText.includes('代理店様用フリーボックス①')) {
                        const tr = span.closest('tr');
                        if (tr) {
                            const td = tr.querySelector('td.dataCol');
                            return td ? td.innerText.trim() : '';
                        }
                    }
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ 三次店の取得に失敗: {e}")
        data["mitsugiten"] = mitsugiten_raw.strip()

        # ── 東西区分 ──
        tozai_raw = _js_get_next_td_text(detail_page, "東西区分")
        if "東日本" in tozai_raw:
            data["tozai"] = "東"
        elif "西日本" in tozai_raw:
            data["tozai"] = "西"
        else:
            data["tozai"] = ""
        self._log(f"  東西区分: {tozai_raw} → {data['tozai']}")

        # ── 取扱店CODE（取次列） ──
        # リンクテキストを含む場合があるのでinnerTextで取得
        toritsugite = _js_get_next_td_text(detail_page, "取扱店CODE")
        data["toritsugite"] = toritsugite
        self._log(f"  取扱店CODE: {toritsugite}")

        # ── NTTパートナーコード（コード列） ──
        # 「NTTパートナーコード判別フラグ」など類似ラベルとの誤マッチを防ぐため完全一致で取得
        code = ""
        try:
            code = detail_page.evaluate("""
            (() => {
                const labels = document.querySelectorAll('td.labelCol');
                for (const td of labels) {
                    if (td.innerText && td.innerText.trim() === 'NTTパートナーコード') {
                        const next = td.nextElementSibling;
                        if (next) return next.innerText.trim();
                    }
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ NTTパートナーコードの取得に失敗: {e}")
        data["code"] = code.strip()
        self._log(f"  NTTパートナーコード: {code}")

        # ── 前確コメント ──
        zenkatsu = ""
        try:
            zenkatsu = detail_page.evaluate("""
            (() => {
                const labels = document.querySelectorAll('td.labelCol');
                for (const td of labels) {
                    if (td.innerText && td.innerText.trim() === '前確コメント') {
                        const next = td.nextElementSibling;
                        return next ? next.innerText.trim() : '';
                    }
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ 前確コメントの取得に失敗: {e}")
        data["zenkatsu_comment"] = zenkatsu.strip()
        self._log(f"  前確コメント: {zenkatsu[:30]}{'...' if len(zenkatsu) > 30 else ''}")

        # ── プラン ──
        plan = ""
        try:
            plan = detail_page.evaluate("""
            (() => {
                const links = document.querySelectorAll("a[href*='/a1i']");
                for (const a of links) {
                    const t = a.innerText.trim();
                    if (t.includes('フレッツ')) return t;
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ プランの取得に失敗: {e}")
        data["plan"] = plan.strip()
        self._log(f"  プラン: {plan}")

        # ── ひかり電話 ──
        hikari_denwa = ""
        try:
            hikari_denwa = detail_page.evaluate("""
            (() => {
                const links = document.querySelectorAll("a[href*='/a1i']");
                for (const a of links) {
                    const t = a.innerText.trim();
                    if (t.includes('ひかり電話') || t.includes('電話')) return t;
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ ひかり電話の取得に失敗: {e}")
        data["hikari_denwa"] = hikari_denwa.strip()
        self._log(f"  ひかり電話: {hikari_denwa}")

        detail_page.close()
        self._log("詳細ページを閉じました")
        time.sleep(0.5)

        return data

    # ══════════════════════════════════════════
    # USEN専用処理
    # ══════════════════════════════════════════

    # USEN検索用ID定数
    ID_USEN_TORITSUGITE_OP  = "j_id0:sve_form1:pageBlock1:searchPageBlockSection1:inputField80_operator"
    ID_USEN_TORITSUGITE_VAL = "j_id0:sve_form1:pageBlock1:searchPageBlockSection1:inputField80"
    ID_USEN_DATE_FROM       = "j_id0:sve_form1:pageBlock1:searchPageBlockSection1:inputSearchFromTo2_from"
    ID_USEN_DATE_TO         = "j_id0:sve_form1:pageBlock1:searchPageBlockSection1:inputSearchFromTo2_to"
    USEN_TORITSUGITE_CODE   = "HCAYVT005"

    def search_usen(self, date_obj) -> int:
        """USEN案件を取扱店コード＋申込日で検索し、件数を返す。"""
        date_str = _date_to_str(date_obj)
        self._log(f"【USEN】取扱店コード検索: {self.USEN_TORITSUGITE_CODE} / 申込日: {date_str}")

        # 「次の文字列と一致する」を選択
        self.bm.page.evaluate(
            f"document.getElementById('{self.ID_USEN_TORITSUGITE_OP}').value = 'eq'"
        )
        self.bm.safe_wait()

        # 取扱店コードを入力
        self.bm.page.evaluate(
            f"document.getElementById('{self.ID_USEN_TORITSUGITE_VAL}').value = '{self.USEN_TORITSUGITE_CODE}'"
        )
        self.bm.safe_wait()

        # 申込日（from・to）を入力
        self.bm.page.evaluate(
            f"document.getElementById('{self.ID_USEN_DATE_FROM}').value = '{date_str}'"
        )
        self.bm.safe_wait()
        self.bm.page.evaluate(
            f"document.getElementById('{self.ID_USEN_DATE_TO}').value = '{date_str}'"
        )
        self.bm.safe_wait()

        self._log("【USEN】検索ボタンをクリック")
        self.bm.page.evaluate(f"document.getElementById('{ID_SEARCH_BTN}').click()")
        self.bm.wait_for_load()

        count_text = self.bm.get_text(SEL_RECORD_COUNT, "件数")
        count = 0
        m = re.search(r"\d+", count_text)
        if m:
            count = int(m.group())
        self._log(f"【USEN】検索結果: {count}件")
        return count

    def get_usen_detail(self, req_url: str) -> dict:
        """
        USEN案件の詳細ページを別タブで開き、
        ①プラン ②ひかり電話 ③法人登録(編集)ページのパートナーコード を取得する。
        東西は「西」固定、取次は「USEN」固定。
        """
        self._log(f"【USEN】詳細ページを別タブで開く: {req_url}")
        context = self.bm._context
        detail_page = context.new_page()
        detail_page.goto(req_url)
        detail_page.wait_for_load_state("domcontentloaded")
        time.sleep(0.5)

        data = {
            "tozai": "西",
            "toritsugite": "USEN",
            "mitsugiten": "",
        }

        # ── プラン ──
        plan = ""
        try:
            plan = detail_page.evaluate("""
            (() => {
                const links = document.querySelectorAll("a[href*='/a1i']");
                for (const a of links) {
                    const t = a.innerText.trim();
                    if (t.includes('フレッツ')) return t;
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ 【USEN】プランの取得に失敗: {e}")
        data["plan"] = plan.strip()
        self._log(f"  【USEN】プラン: {plan}")

        # ── ひかり電話 ──
        hikari_denwa = ""
        try:
            hikari_denwa = detail_page.evaluate("""
            (() => {
                const links = document.querySelectorAll("a[href*='/a1i']");
                for (const a of links) {
                    const t = a.innerText.trim();
                    if (t.includes('ひかり電話') || t.includes('電話')) return t;
                }
                return '';
            })()
            """) or ""
        except Exception as e:
            self._log(f"⚠ 【USEN】ひかり電話の取得に失敗: {e}")
        data["hikari_denwa"] = hikari_denwa.strip()
        self._log(f"  【USEN】ひかり電話: {hikari_denwa}")

        # ── ⑥ 法人登録（編集）ボタンを押して別タブでパートナーコードを取得 ──
        code = ""
        try:
            # ボタンのonclick属性からURLを抽出して直接遷移（新タブ制御が複雑なため）
            entry_url = detail_page.evaluate("""
            (() => {
                const btns = document.querySelectorAll('input[name="entry_edit_co"]');
                if (btns.length === 0) return '';
                const onclick = btns[0].getAttribute('onclick') || '';
                // navigateToUrl('URL','DETAIL') の形式からURL部分を抽出
                const m = onclick.match(/navigateToUrl\\('([^']+)'/);
                return m ? m[1] : '';
            })()
            """) or ""

            if entry_url:
                self._log(f"  【USEN】法人登録(編集)ページへ移動: {entry_url}")
                entry_page = context.new_page()
                entry_page.goto(entry_url)
                entry_page.wait_for_load_state("domcontentloaded")
                time.sleep(0.5)

                # ⑦ NTTパートナーコード（手入力）テキストボックスの値を取得
                code = entry_page.evaluate("""
                (() => {
                    // idの末尾がComponent149のinputを直接取得
                    const inputs = document.querySelectorAll('input[type="text"]');
                    for (const inp of inputs) {
                        if ((inp.id || '').endsWith('Component149')) {
                            return inp.value.trim();
                        }
                    }
                    // フォールバック: 10桁数値のinputを探す
                    for (const inp of inputs) {
                        if (/^\d{10}$/.test((inp.value || '').trim())) {
                            return inp.value.trim();
                        }
                    }
                    return '';
                })()
                """) or ""
                self._log(f"  【USEN】NTTパートナーコード: {code}")

                # ── 代理店様用フリーボックス① (inputField123) ──
                mitsugiten_raw = ""
                try:
                    mitsugiten_raw = entry_page.eval_on_selector(
                        "input[id$='inputField123']",
                        "el => el.value.trim()"
                    ) or ""
                except Exception:
                    pass
                data["mitsugiten"] = mitsugiten_raw.strip()
                self._log(f"  【USEN】三次店(raw): {mitsugiten_raw}")

                # ⑧ タブを閉じる
                entry_page.close()
                time.sleep(0.5)
            else:
                self._log("⚠ 【USEN】法人登録(編集)ボタンが見つかりませんでした")

        except Exception as e:
            self._log(f"⚠ 【USEN】パートナーコード取得に失敗: {e}")

        data["code"] = code.strip()

        detail_page.close()
        self._log("【USEN】詳細ページを閉じました")
        time.sleep(0.5)

        return data
