# YukarinetteLogger

YukarinetteLogger は、ゆかりねっとコネクター Neo から送られてくる **翻訳テキスト（ja/en）を WebSocket 経由で受信し、ログとして保存する常駐アプリ**です。タスクトレイに常駐し、WebSocket の接続状態をリアルタイムに表示します。翻訳ログを後から確認したい配信者・実況者向けの補助ツールです。

---

## ✨ 主な機能

### ✔ WebSocket 受信
- ゆかりねっとコネクター Neo からの翻訳データを WebSocket で受信
- レジストリから WebSocket のポート番号を自動取得（DWORD の場合は `ws://127.0.0.1:PORT` として組み立て）

### ✔ MessageID ベースのログ確定処理
- 同じ MessageID のデータは上書きし、最後の1件だけをログに保存
- MessageID が変わったら前のデータを確定出力
- 一定時間更新がなければ自動で確定出力

### ✔ ログ出力形式
```
YYYYMMDD-HH:MM:SSSSS ja:日本語,en:English
```

### ✔ exe 名に合わせたログファイル名
- YukarinetteLogger.exe → YukarinetteLogger.log
- Python スクリプトとして実行した場合は main.log

### ✔ タスクトレイ常駐
- WebSocket の状態をツールチップで表示  
  - 接続待機中  
  - 接続中  
  - 受信中  
  - 再接続中  
  - 切断  
- 右クリックメニューに Exit（終了）を表示

### ✔ Windows 通知
- 起動時  
- WebSocket 接続成功時  
- 再接続失敗時  

---

## 📦 インストール方法

### 1. Release から ZIP をダウンロード
```
YukarinetteLogger.zip
├─ YukarinetteLogger.exe
└─ config.json
```

### 2. 任意のフォルダに展開

### 3. config.json を編集（必要に応じて）
```json
{
  "DEBUG": false,
  "REGISTRY_HIVE": "HKEY_CURRENT_USER",
  "REGISTRY_PATH": "SOFTWARE\\YukarinetteConnectorNeo",
  "REGISTRY_VALUE": "WebSocketPort",
  "LOG_DIR": "log",
  "PROCESS_STABLE_SEC": 10,
  "FLUSH_INTERVAL_SEC": 5,
  "WS_RECONNECT_DELAY_SEC": 5,
  "WS_MAX_RECONNECT_SEC": 60
}
```

### 4. YukarinetteLogger.exe を起動  
タスクトレイにアイコンが表示されます。

---

## 📝 レジストリ設定について

WebSocket のポート番号はレジストリから取得します。

### ✔ REG_SZ（文字列）の場合
そのまま URL として使用します。
```
ws://127.0.0.1:12345
```

### ✔ DWORD（数値）の場合
ポート番号として扱い、以下の URL を自動生成します。
```
ws://127.0.0.1:{PORT}
```

---

## 🛠 ビルド方法（開発者向け）

### 1. 依存関係をインストール
```
pip install -r requirements.txt
```

### 2. PyInstaller でビルド
```
pyinstaller --onefile --noconsole --name YukarinetteLogger main.py
```

### 3. config.json を dist フォルダへコピー
```
copy config.json dist\
```

---

## 🚀 GitHub Actions（自動ビルド & Release）

- タグ（例：v1.0.0）を push すると自動ビルド  
- YukarinetteLogger.exe と config.json を ZIP 化  
- YukarinetteLogger.zip として Release にアップロード  

---

## 📄 ライセンス

MIT License（必要に応じて変更してください）

---

## 🙏 作者より

このツールは、ゆかりねっとコネクター Neo をより便利に使うための **翻訳ログ保存専用アプリ**として開発されました。改善案・要望・バグ報告などあれば Issue へどうぞ。
