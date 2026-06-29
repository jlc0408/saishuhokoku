"""
フレッツ詳細エントリーアプリ メインUI

【設計上の重要な制約】
Playwright sync_api はスレッドをまたいで使用できない。
そのため Edge起動ボタンは「起動済みフラグを立てるだけ」にし、
スタートボタンが押されたときに初めてワーカースレッドを起動して
そのスレッド内で launch() → login() → 処理 をすべて行う。
"""
from __future__ import annotations
import os
import threading
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import config_manager

WINDOW_W = 1280
WINDOW_H = 720

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
    _BaseApp = TkinterDnD.Tk
except Exception:
    _DND_AVAILABLE = False
    DND_FILES = None
    _BaseApp = tk.Tk


class App(_BaseApp):
    def __init__(self):
        super().__init__()
        self.title("最終報告")
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")

        # アイコン設定（タスクバー・タイトルバー・タブ）
        _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
        if os.path.exists(_icon_path):
            self.iconbitmap(default=_icon_path)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self._excel_path: str = ""
        self._selected_date: datetime.date | None = None
        self._running = False
        self._bm = None  # BrowserManager（ワーカースレッド内でのみ使用）

        self._build_ui()

    # ─────────────────────────────────────────
    # UI 構築
    # ─────────────────────────────────────────
    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self._main_tab = ttk.Frame(notebook)
        notebook.add(self._main_tab, text="メイン操作")
        self._build_main_tab(self._main_tab)

        self._config_tab = ttk.Frame(notebook)
        notebook.add(self._config_tab, text="設定")
        self._build_config_tab(self._config_tab)

    def _build_main_tab(self, parent: ttk.Frame):
        pad = {"padx": 12, "pady": 6}

        # ─ ① 日付設定 ─
        date_frame = ttk.LabelFrame(parent, text="① 取得する日付")
        date_frame.pack(fill="x", **pad)

        self._date_var = tk.StringVar(value="today")

        self._btn_today = tk.Button(
            date_frame, text="本日", width=6,
            relief="raised", bg="#e0e0e0", fg="black", font=("MS Gothic", 10),
            command=lambda: self._select_date_mode("today"),
        )
        self._btn_today.pack(side="left", padx=(8, 4), pady=6)

        self._btn_yesterday = tk.Button(
            date_frame, text="昨日", width=6,
            relief="raised", bg="#e0e0e0", fg="black", font=("MS Gothic", 10),
            command=lambda: self._select_date_mode("yesterday"),
        )
        self._btn_yesterday.pack(side="left", padx=4, pady=6)

        ttk.Label(date_frame, text="開始:").pack(side="left", padx=(12, 2))
        self._date_from_entry = ttk.Entry(date_frame, width=12)
        self._date_from_entry.pack(side="left")

        ttk.Label(date_frame, text="終了:").pack(side="left", padx=(8, 2))
        self._date_to_entry = ttk.Entry(date_frame, width=12)
        self._date_to_entry.pack(side="left")

        ttk.Label(date_frame, text="(YYYY/MM/DD)").pack(side="left", padx=(4, 0))

        self._update_date_label()

        # ─ ③ エクセル ─
        excel_frame = ttk.LabelFrame(parent, text="② Excelファイル")
        excel_frame.pack(fill="x", **pad)

        ttk.Button(excel_frame, text="ファイルを選択…",
                   command=self._open_excel_dialog).pack(side="left", padx=8, pady=6)

        self._excel_label = ttk.Label(excel_frame, text="← ボタンから選択、またはウィンドウにドロップ", foreground="gray")
        self._excel_label.pack(side="left", padx=4)

        # ─ ウィンドウ全体をドロップターゲットに ─
        if _DND_AVAILABLE:
            parent.drop_target_register(DND_FILES)
            parent.dnd_bind("<<Drop>>", self._on_drop)

        # ─ ④ スタート ─
        start_frame = ttk.Frame(parent)
        start_frame.pack(fill="x", **pad)

        self._btn_start = ttk.Button(start_frame, text="▶  スタート",
                                     command=self._start)
        self._btn_start.pack(side="left", padx=8, pady=6, ipadx=20, ipady=6)

        self._btn_stop = ttk.Button(start_frame, text="■  停止",
                                    command=self._stop, state="disabled")
        self._btn_stop.pack(side="left", padx=4)

        self._status_label = ttk.Label(start_frame, text="待機中", foreground="gray")
        self._status_label.pack(side="left", padx=16)

        # ─ ログ ─
        log_frame = ttk.LabelFrame(parent, text="ログ")
        log_frame.pack(fill="both", expand=True, **pad)

        self._log_text = tk.Text(log_frame, state="disabled", wrap="word",
                                 font=("MS Gothic", 10))
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical",
                                  command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._log_text.pack(fill="both", expand=True)

        ttk.Button(parent, text="ログをコピー",
                   command=self._copy_log).pack(anchor="e", padx=12, pady=2)

    def _build_config_tab(self, parent: ttk.Frame):
        pad = {"padx": 16, "pady": 8}

        cfg_frame = ttk.LabelFrame(parent, text="BEAMS ログイン設定")
        cfg_frame.pack(fill="x", **pad)

        ttk.Label(cfg_frame, text="ログインURL:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self._url_var = tk.StringVar(value=config_manager.get("BEAMS", "url"))
        ttk.Entry(cfg_frame, textvariable=self._url_var, width=60).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(cfg_frame, text="ユーザーID:").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self._user_var = tk.StringVar(value=config_manager.get("BEAMS", "username"))
        ttk.Entry(cfg_frame, textvariable=self._user_var, width=40).grid(row=1, column=1, sticky="w")

        ttk.Label(cfg_frame, text="パスワード:").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        self._pass_var = tk.StringVar(value=config_manager.get("BEAMS", "password"))
        ttk.Entry(cfg_frame, textvariable=self._pass_var, show="*", width=40).grid(row=2, column=1, sticky="w")

        edge_frame = ttk.LabelFrame(parent, text="Edge パス設定")
        edge_frame.pack(fill="x", **pad)

        ttk.Label(edge_frame, text="msedge.exe パス:").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self._edge_path_var = tk.StringVar(value=config_manager.get("EDGE", "edge_path"))
        ttk.Entry(edge_frame, textvariable=self._edge_path_var, width=60).grid(row=0, column=1, sticky="w")
        ttk.Button(edge_frame, text="参照…",
                   command=self._browse_edge).grid(row=0, column=2, padx=4)

        ttk.Button(parent, text="設定を保存", command=self._save_config).pack(**pad)

    # ─────────────────────────────────────────
    # 日付処理
    # ─────────────────────────────────────────
    def _select_date_mode(self, mode: str):
        """ボタンクリック時に日付をセットし、150ms後にニュートラルに戻す。"""
        self._date_var.set(mode)
        btn = self._btn_today if mode == "today" else self._btn_yesterday
        btn.configure(relief="sunken", bg="#4a90d9", fg="white")
        self._update_date_label()
        self.after(150, lambda: btn.configure(relief="raised", bg="#e0e0e0", fg="black"))

    def _update_date_label(self):
        """本日・昨日ボタン押下時に開始日・終了日を同じ値でセットする。"""
        mode = self._date_var.get()
        today = datetime.date.today()
        d = today if mode == "today" else today - datetime.timedelta(days=1)
        ds = d.strftime("%Y/%m/%d")
        for entry in (self._date_from_entry, self._date_to_entry):
            entry.delete(0, "end")
            entry.insert(0, ds)

    def _get_date_range(self) -> tuple[datetime.date, datetime.date] | tuple[None, None]:
        """開始日・終了日をパースして返す。不正な場合は (None, None)。"""
        try:
            d_from = datetime.datetime.strptime(self._date_from_entry.get().strip(), "%Y/%m/%d").date()
            d_to   = datetime.datetime.strptime(self._date_to_entry.get().strip(), "%Y/%m/%d").date()
            return d_from, d_to
        except ValueError:
            return None, None

    # ─────────────────────────────────────────
    # ファイル選択
    # ─────────────────────────────────────────
    def _open_excel_dialog(self):
        path = filedialog.askopenfilename(
            title="Excelファイルを選択",
            filetypes=[("Excelファイル", "*.xlsx *.xlsm"), ("すべてのファイル", "*.*")],
        )
        if path:
            self._set_excel(path)

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        if path.lower().endswith((".xlsx", ".xlsm")):
            self._set_excel(path)
        else:
            self._log("⚠ xlsx または xlsm ファイルをドロップしてください。")

    def _set_excel(self, path: str):
        self._excel_path = path
        self._excel_label.configure(text=os.path.basename(path), foreground="black")
        self._log(f"Excelファイルを選択: {path}")

    # ─────────────────────────────────────────
    # スタート（Edge起動 → 処理 を同一スレッドで実行）
    # ─────────────────────────────────────────
    def _start(self):
        # ① ID・パス未設定チェック
        uid = config_manager.get("BEAMS", "username").strip()
        pwd = config_manager.get("BEAMS", "password").strip()
        if not uid or not pwd:
            messagebox.showerror(
                "設定が必要です",
                "BEAMSのユーザーIDまたはパスワードが設定されていません。\n"
                "「設定」タブでIDとパスワードを入力して保存してください。"
            )
            return

        date_from, date_to = self._get_date_range()
        if not date_from or not date_to:
            messagebox.showerror("エラー", "日付を正しく設定してください（例: 2026/06/02）")
            return
        if date_from > date_to:
            messagebox.showerror("エラー", "開始日は終了日以前の日付を設定してください。")
            return
        if not self._excel_path:
            messagebox.showerror("エラー", "Excelファイルを選択してください。")
            return

        self._running = True
        self._btn_start.configure(state="disabled")
        self._btn_stop.configure(state="normal")
        self._set_status("Edge起動中…", "orange")
        self._log("═" * 50)
        self._log(f"処理開始 | {date_from.strftime('%Y/%m/%d')} ～ {date_to.strftime('%Y/%m/%d')} | ファイル: {os.path.basename(self._excel_path)}")

        threading.Thread(target=self._run_all, args=(date_from, date_to), daemon=True).start()

    def _stop(self):
        self._running = False
        self._log("⚠ ユーザーにより停止要求")
        self._btn_stop.configure(state="disabled")

    def _run_all(self, date_from: datetime.date, date_to: datetime.date):
        """Edge起動 → ログイン → 処理 をすべてこのスレッド内で実行する。"""
        from browser_manager import BrowserManager, BrowserError
        from beams_scraper import BeamsScraper
        from excel_manager import ExcelManager

        excel: ExcelManager | None = None
        bm: BrowserManager | None = None

        try:
            # ① Edge起動
            bm = BrowserManager(logger=self._log)
            bm.launch()
            self._bm = bm  # 停止ボタン用に保持
            self._set_status("Edge起動済み", "green")

            if not self._running:
                return

            # エクセルを開く
            self._log("Excelファイルを読み込み中…")
            try:
                excel = ExcelManager(self._excel_path)
            except (PermissionError, OSError) as e:
                msg = (
                    f"Excelファイルを開けませんでした。\n\n"
                    f"【原因】ファイルが既に Excel で開かれている可能性があります。\n"
                    f"【対処】Excel を閉じてから、もう一度スタートしてください。\n\n"
                    f"ファイル: {self._excel_path}\n"
                    f"エラー詳細: {e}"
                )
                self._log(f"❌ {msg}")
                self.after(0, lambda m=msg: messagebox.showerror("Excelファイルを開けません", m))
                return

            scraper = BeamsScraper(bm, logger=self._log)
            self._set_status("処理中…", "blue")

            # ⑤ ログイン
            scraper.login()
            if not self._running:
                return

            # ⑥ 申込検索Bタブ
            scraper.go_to_search_tab()
            if not self._running:
                return

            # ⑦⑧ 検索
            count = scraper.search_by_date(date_from, date_to)
            if not self._running:
                return

            # ⑨ 件数チェック
            if count > 50:
                self._log(f"⚠ 検索結果が{count}件あります（51件以上）。全件を処理します。")

            # ⑩ REQリスト取得
            req_list = scraper.get_req_list()
            if not req_list:
                self._log("処理対象の案件が見つかりませんでした。")
                return

            et_date_str = self._date_to_et(date_from)  # ET日は開始日を使用
            success = 0
            fail = 0

            for i, req_item in enumerate(req_list):
                if not self._running:
                    self._log("処理を中断しました。")
                    break

                req_no = req_item["req"]
                req_url = req_item["url"]
                self._log(f"[{i+1}/{len(req_list)}] {req_no} を処理中…")
                self._set_status(f"処理中 {i+1}/{len(req_list)}", "blue")

                try:
                    detail = scraper.get_detail(req_url)
                    row = excel.find_first_empty_row()
                    record = {"req": req_no, "req_url": req_url, "et_date": et_date_str, **detail}
                    converted = excel.write_record(row, record)
                    self._log(
                        f"  ✔ {req_no} → 行{row} | 三次店:{converted.get('mitsugiten','')} | "
                        f"東西:{converted.get('tozai','')} | 取次:{converted.get('toritsugite','')} | "
                        f"コード:{converted.get('code','')} | プラン:{converted.get('plan','')}"
                    )
                    success += 1
                except BrowserError as e:
                    self._log(f"  ❌ {req_no} エラー: {e}")
                    fail += 1
                except Exception as e:
                    self._log(f"  ❌ {req_no} 予期しないエラー: {e}")
                    fail += 1

            # ──────────────────────────────────────────
            # USEN処理
            # ──────────────────────────────────────────
            self._log("═" * 50)
            self._log("【USEN】処理開始")
            self._set_status("USEN処理中…", "blue")

            # 申込検索Bタブに戻って検索条件をリセット（再移動）
            scraper.go_to_search_tab()
            usen_count = scraper.search_usen(date_from, date_to)

            if usen_count > 50:
                self._log(f"⚠ 【USEN】検索結果が{usen_count}件あります（51件以上）。全件を処理します。")

            usen_req_list = scraper.get_req_list()
            self._log(f"【USEN】{len(usen_req_list)}件のREQを取得")

            usen_success = 0
            usen_fail = 0

            for i, req_item in enumerate(usen_req_list):
                if not self._running:
                    self._log("処理を中断しました。")
                    break

                req_no = req_item["req"]
                req_url = req_item["url"]
                self._log(f"【USEN】[{i+1}/{len(usen_req_list)}] {req_no} を処理中…")
                self._set_status(f"USEN処理中 {i+1}/{len(usen_req_list)}", "blue")

                try:
                    detail = scraper.get_usen_detail(req_url)
                    row = excel.find_first_empty_row()
                    record = {"req": req_no, "req_url": req_url, "et_date": et_date_str, **detail}
                    converted = excel.write_record(row, record)
                    self._log(
                        f"  ✔ {req_no} → 行{row} | 東西:{converted.get('tozai','')} | "
                        f"取次:{converted.get('toritsugite','')} | "
                        f"コード:{converted.get('code','')} | プラン:{converted.get('plan','')}"
                    )
                    usen_success += 1
                except BrowserError as e:
                    self._log(f"  ❌ {req_no} エラー: {e}")
                    usen_fail += 1
                except Exception as e:
                    self._log(f"  ❌ {req_no} 予期しないエラー: {e}")
                    usen_fail += 1

            total_success = success + usen_success
            total_fail = fail + usen_fail

            excel.save()
            self._log("═" * 50)
            self._log(f"フレッツ 成功:{success}件 失敗:{fail}件 / USEN 成功:{usen_success}件 失敗:{usen_fail}件")
            self._log(f"Excelを保存しました: {self._excel_path}")

            if total_fail > 0:
                self.after(0, lambda: messagebox.showwarning(
                    "処理完了（一部エラーあり）",
                    f"フレッツ 成功:{success}件 失敗:{fail}件\n"
                    f"USEN  成功:{usen_success}件 失敗:{usen_fail}件\n"
                    f"ログを確認してください。"
                ))
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "処理完了",
                    f"フレッツ {success}件 / USEN {usen_success}件\n全件正常に処理しました。"
                ))

        except Exception as e:
            self._log(f"❌ 致命的エラー: {e}")
            self.after(0, lambda: messagebox.showerror("エラー", str(e)))
        finally:
            if excel:
                excel.close()
            if self._bm:
                self._bm.close()
                self._bm = None
            self.after(0, self._on_automation_done)

    def _on_automation_done(self):
        self._running = False
        self._btn_start.configure(state="normal")
        self._btn_stop.configure(state="disabled")
        self._set_status("待機中", "gray")

    # ─────────────────────────────────────────
    # ログ・ステータス
    # ─────────────────────────────────────────
    def _log(self, msg: str):
        def _append():
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self._log_text.configure(state="normal")
            self._log_text.insert("end", f"[{ts}] {msg}\n")
            self._log_text.configure(state="disabled")
            self._log_text.see("end")
        self.after(0, _append)

    def _set_status(self, text: str, color: str):
        self.after(0, lambda: self._status_label.configure(text=text, foreground=color))

    def _copy_log(self):
        content = self._log_text.get("1.0", "end")
        self.clipboard_clear()
        self.clipboard_append(content)
        self._log("ログをクリップボードにコピーしました。")

    # ─────────────────────────────────────────
    # 設定
    # ─────────────────────────────────────────
    def _save_config(self):
        config_manager.set_value("BEAMS", "url", self._url_var.get())
        config_manager.set_value("BEAMS", "username", self._user_var.get())
        config_manager.set_value("BEAMS", "password", self._pass_var.get())
        config_manager.set_value("EDGE", "edge_path", self._edge_path_var.get())
        messagebox.showinfo("保存完了", "設定を保存しました。")

    def _browse_edge(self):
        path = filedialog.askopenfilename(
            title="msedge.exe を選択",
            filetypes=[("Edge 実行ファイル", "msedge.exe"), ("すべて", "*.*")],
        )
        if path:
            self._edge_path_var.set(path)

    @staticmethod
    def _date_to_et(d: datetime.date) -> str:
        return f"{d.month}月{d.day}日"

    def _on_close(self):
        if self._running:
            if not messagebox.askyesno("確認", "処理中です。終了しますか？"):
                return
        if self._bm:
            self._bm.close()
        self.destroy()


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
