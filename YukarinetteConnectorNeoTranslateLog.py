import asyncio
import json
import logging
import os
import sys
import threading
import time
from datetime import datetime

import winreg
import websockets
from PIL import Image, ImageDraw
from win10toast import ToastNotifier
import pystray


# ==============================
# パス関連（実行ディレクトリの扱い）
# ==============================

def get_program_dir():
    """
    実行ファイルが置いてあるフォルダを返す。
    PyInstaller で exe 化しても、.py 実行でも期待どおりになる。
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller などで固めた場合
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        # 通常の .py 実行の場合
        return os.path.dirname(os.path.abspath(__file__))


PROGRAM_DIR = get_program_dir()


# ==============================
# config 読み込み
# ==============================

def load_config():
    path = os.path.join(PROGRAM_DIR, "config.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


config = load_config()


# ==============================
# ログ設定
# ==============================

log_level = logging.DEBUG if config.get("DEBUG", False) else logging.INFO

log_file_path = os.path.join(PROGRAM_DIR, "app.log")
logging.basicConfig(
    level=log_level,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

logging.info("=== Application started ===")
logging.debug(f"Program directory: {PROGRAM_DIR}")


# ==============================
# Windows 通知
# ==============================

notifier = ToastNotifier()


def notify(title: str, msg: str):
    try:
        notifier.show_toast(title, msg, duration=3, threaded=True)
    except Exception as e:
        logging.error(f"Notify error: {e}")


# ==============================
# レジストリ読み取り
# ==============================

HIVE_MAP = {
    "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
    "HKLM": winreg.HKEY_LOCAL_MACHINE,
    "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
    "HKCU": winreg.HKEY_CURRENT_USER,
}


def read_registry_value():
    hive_name = config.get("REGISTRY_HIVE", "HKEY_LOCAL_MACHINE")
    hive = HIVE_MAP.get(hive_name.upper(), winreg.HKEY_LOCAL_MACHINE)

    path = config["REGISTRY_PATH"]
    value_name = config["REGISTRY_VALUE"]

    logging.debug(f"Reading registry: {hive_name}\\{path} ({value_name})")

    try:
        key = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(key, value_name)
        logging.info(f"Registry value: {value}")
        return value
    except Exception as e:
        logging.error(f"Registry read error: {e}")
        return None


# ==============================
# JSON ログ出力（処理ID管理付き）
# ==============================

LOG_DIR = os.path.join(PROGRAM_DIR, config.get("LOG_DIR", "log"))
os.makedirs(LOG_DIR, exist_ok=True)

PROCESS_ID_KEY = config.get("PROCESS_ID_KEY", "process_id")
PROCESS_STABLE_SEC = config.get("PROCESS_STABLE_SEC", 10)
FLUSH_INTERVAL_SEC = config.get("FLUSH_INTERVAL_SEC", 5)


class ProcessBuffer:
    """
    同じ処理IDのデータをバッファし、
    ・処理ID変更時
    ・一定時間（PROCESS_STABLE_SEC）経過時
    に最後の1件だけをログ出力する
    """

    def __init__(self):
        self.current_id = None
        self.last_data = None
        self.last_update_time = None
        self.lock = asyncio.Lock()

    async def add_message(self, raw_data: str):
        """
        WebSocket から受信した JSON 文字列を処理
        """
        try:
            data = json.loads(raw_data)
        except json.JSONDecodeError as e:
            logging.error(f"JSON decode error: {e} / raw={raw_data}")
            return

        proc_id = data.get(PROCESS_ID_KEY)
        if proc_id is None:
            logging.warning(f"PROCESS_ID_KEY '{PROCESS_ID_KEY}' not found in data: {data}")
            return

        async with self.lock:
            now = time.time()

            if self.current_id is None:
                # 初回
                self.current_id = proc_id
                self.last_data = data
                self.last_update_time = now
                logging.debug(f"New process_id={proc_id} started")
                return

            if proc_id != self.current_id:
                # 別の処理IDが来た → 直前の処理IDの最後のデータを確定出力
                logging.debug(f"Process_id changed: {self.current_id} -> {proc_id}")
                await self._flush_locked()
                # 新しい処理IDとしてセット
                self.current_id = proc_id
                self.last_data = data
                self.last_update_time = now
            else:
                # 同じ処理ID → データを上書き（最後の1件を覚える）
                self.last_data = data
                self.last_update_time = now
                logging.debug(f"Updated data for process_id={proc_id}")

    async def periodic_flush(self):
        """
        一定間隔で呼び出して、
        「一定時間更新がない処理ID」を確定として出力する
        """
        while True:
            try:
                await asyncio.sleep(FLUSH_INTERVAL_SEC)
                async with self.lock:
                    if self.current_id is None or self.last_update_time is None:
                        continue
                    now = time.time()
                    if now - self.last_update_time >= PROCESS_STABLE_SEC:
                        logging.debug(
                            f"Process_id={self.current_id} became stable "
                            f"({now - self.last_update_time:.1f}s >= {PROCESS_STABLE_SEC}s)"
                        )
                        await self._flush_locked()
            except asyncio.CancelledError:
                break
            except Exception as e:
                logging.error(f"periodic_flush error: {e}")

    async def _flush_locked(self):
        """
        現在の last_data をログ出力して、バッファをクリア
        呼び出し元で lock を取っている前提
        """
        if self.current_id is None or self.last_data is None:
            return

        # ファイル名: timestamp_processid.json
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        safe_id = str(self.current_id).replace(os.sep, "_")
        filename = f"{timestamp}_{safe_id}.json"
        path = os.path.join(LOG_DIR, filename)

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.last_data, f, ensure_ascii=False, indent=2)
            logging.info(f"Saved JSON: {path}")
        except Exception as e:
            logging.error(f"JSON save error: {e}")

        # クリア
        self.current_id = None
        self.last_data = None
        self.last_update_time = None


process_buffer = ProcessBuffer()


# ==============================
# WebSocket 受信 + 再接続ロジック
# ==============================

WS_RECONNECT_DELAY_SEC = config.get("WS_RECONNECT_DELAY_SEC", 5)
WS_MAX_RECONNECT_SEC = config.get("WS_MAX_RECONNECT_SEC", 60)


async def websocket_loop(url: str):
    """
    WebSocket に接続し、切断されたらリトライ。
    一定時間再接続できなければアプリ終了。
    """
    notify("起動", "アプリケーションを起動しました")

    start_retry_time = None

    while True:
        try:
            logging.info(f"Connecting WebSocket: {url}")
            async with websockets.connect(url) as ws:
                notify("データ受信準備完了", "WebSocket 接続が確立しました")
                logging.info("WebSocket connected")

                # 再接続タイマをリセット
                start_retry_time = None

                # メイン受信ループ
                async for message in ws:
                    logging.debug(f"Received: {message}")
                    await process_buffer.add_message(message)

        except asyncio.CancelledError:
            logging.info("WebSocket loop cancelled")
            break

        except Exception as e:
            logging.error(f"WebSocket error/disconnected: {e}")

            # 初回切断時刻を記録
            if start_retry_time is None:
                start_retry_time = time.time()

            elapsed = time.time() - start_retry_time
            if elapsed >= WS_MAX_RECONNECT_SEC:
                logging.error("WebSocket reconnect timeout exceeded. Exiting.")
                notify("終了", "WebSocket 再接続に失敗したため終了します")
                # 残っているデータを確定出力しておきたいならここで flush を呼ぶ
                async with process_buffer.lock:
                    await process_buffer._flush_locked()
                os._exit(1)

            logging.info(f"Retrying WebSocket in {WS_RECONNECT_DELAY_SEC} sec...")
            await asyncio.sleep(WS_RECONNECT_DELAY_SEC)


# ==============================
# タスクトレイアイコン
# ==============================

def create_icon_image():
    img = Image.new("RGB", (16, 16), "blue")
    d = ImageDraw.Draw(img)
    d.rectangle([4, 4, 12, 12], fill="white")
    return img


def on_exit(icon, item):
    notify("終了", "アプリケーションを終了します")
    icon.stop()
    os._exit(0)


def run_tray():
    icon = pystray.Icon(
        "MyTrayApp",
        create_icon_image(),
        "My Python Tray App",
        menu=pystray.Menu(
            pystray.MenuItem("Exit", on_exit),
        ),
    )
    icon.run()


# ==============================
# メイン
# ==============================

async def main_async():
    # レジストリから WebSocket URL を取得
    ws_url = read_registry_value()
    if not ws_url:
        logging.error("WebSocket URL がレジストリから取得できませんでした。終了します。")
        notify("終了", "設定取得エラーのため終了します")
        return

    # 処理IDの「一定時間経過による確定出力」タスク
    flush_task = asyncio.create_task(process_buffer.periodic_flush())
    ws_task = asyncio.create_task(websocket_loop(ws_url))

    try:
        await ws_task
    finally:
        flush_task.cancel()
        try:
            await flush_task
        except asyncio.CancelledError:
            pass
        logging.info("main_async finished")


def main():
    # タスクトレイを別スレッドで起動
    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()

    # asyncio メインループ
    try:
        asyncio.run(main_async())
    except KeyboardInterrupt:
        logging.info("KeyboardInterrupt - exiting")
    finally:
        notify("終了", "アプリケーションを終了します")
        logging.info("=== Application exited ===")


if __name__ == "__main__":
    main()
