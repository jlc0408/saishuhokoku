# フレッツ詳細エントリーアプリ

## ファイル構成

```
flets_entry/
├── main.py              # メインUI（起動ファイル）
├── config_manager.py    # config.ini 読み書き
├── excel_manager.py     # Excel 操作
├── browser_manager.py   # Edge 起動・Playwright 共通制御
├── beams_scraper.py     # BEAMS 画面操作・データ取得
├── requirements.txt     # 依存ライブラリ
└── config.ini           # 設定ファイル（初回起動時に自動生成）
```

## セットアップ

```bash
pip install -r requirements.txt
playwright install chromium
```

## 起動方法

```bash
python main.py
```

## 操作手順

1. **設定タブ** でユーザーID・パスワード・Edge パスを入力して「設定を保存」
2. **① Edge を起動する** ボタンをクリック
3. **② 日付** を選択（本日 / 昨日 / その他）
4. **③ エクセルファイル** を選択またはドロップ
5. **▶ スタート** をクリック → 自動処理開始

## エラーコード

| コード   | 内容                                     |
|---------|------------------------------------------|
| ERR-001 | 指定要素が4回連続で見つからなかった         |

## 注意事項

- エクセルの「フレッツ詳細」シートに以下のヘッダーが1行目に必要です:
  `REQ`, `ET日`, `三次店`, `東西区分`, `取次`, `コード`, `プラン`, `ひかり電話`
- 検索結果が51件以上の場合、処理後にアラートが表示されます
- Undo/Redo は使用しません（禁止ブラウザ対応）
