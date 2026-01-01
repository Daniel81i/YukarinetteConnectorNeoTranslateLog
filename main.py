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
import psutil

tray_icon = None
ws = None

# ==============================
# 実行ファイルのあるディレクトリ取得
# ==============================
def get_program_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))


PROGRAM_DIR = get_program_dir()


# ==============================
# config.json 読み込み
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

# ==============================
# exe 名からログファイル名を決定
# ==============================
def get_exe_name():
    if getattr(sys, 'frozen', False):
        # PyInstaller で exe 化されている場合
        return os.path.splitext(os.path.basename(sys.executable))[0]
    else:
        # Python スクリプトとして実行されている場合
        return os.path.splitext(os.path.basename(__file__))[0]

APP_NAME = get_exe_name()

# ==============================
# タイムスタンプ付きログファイル名
# ==============================
timestamp = datetime.now().strftime("%Y-%m-%d-%H%M%S")
LOG_FILENAME = f"{get_exe_name()}-{timestamp}.log"
LOG_PATH = os.path.join(PROGRAM_DIR, LOG_FILENAME)

logging.basicConfig(
    level=log_level,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)

logging.info("=== Application started ===")


# ==============================
# Windows 通知
# ==============================
notifier = ToastNotifier()


def notify(title, msg):
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

    try:
        key = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
        value, _ = winreg.QueryValueEx(key, value_name)
        logging.info(f"Registry value: {value}")
        return value
    except Exception as e:
        logging.error(f"Registry read error: {e}")
        return None


# ==============================
# MsgID ベースのログ確定処理
# ==============================
LOG_DIR = os.path.join(PROGRAM_DIR, config.get("LOG_DIR", "log"))
os.makedirs(LOG_DIR, exist_ok=True)


class MessageBuffer:
    def __init__(self, stable_sec=10, flush_interval=5):
        self.current_id = None
        self.last_data = None
        self.last_update_time = None
        self.stable_sec = stable_sec
        self.flush_interval = flush_interval
        self.lock = asyncio.Lock()

    async def add_message(self, raw_data: str):
        try:
            data = json.loads(raw_data)
            if logging.getLogger().isEnabledFor(logging.DEBUG):
                logging.debug(f"[WS RECV JSON] {data}")
        
        except json.JSONDecodeError:
            logging.error(f"JSON decode error: {raw_data}")
            return

        msg_id = data.get("MsgID")
        if msg_id is None:
            logging.warning("MsgID が存在しないデータを受信")
            return

        async with self.lock:
            now = time.time()

            if self.current_id is None:
                self.current_id = msg_id
                self.last_data = data
                self.last_update_time = now
                return

            if msg_id != self.current_id:
                await self._flush_locked()
                self.current_id = msg_id
                self.last_data = data
                self.last_update_time = now
            else:
                self.last_data = data
                self.last_update_time = now

    async def periodic_flush(self):
        while True:
            await asyncio.sleep(self.flush_interval)
            async with self.lock:
                if self.current_id is None:
                    continue
                now = time.time()
                if now - self.last_update_time >= self.stable_sec:
                    await self._flush_locked()

    async def _flush_locked(self):
        if self.last_data is None:
            return

        timestamp = datetime.now().strftime("%Y%m%d-%H:%M:%S%f")[:-3]

        lang1 = self.last_data.get("Lang1", "")
        lang2 = self.last_data.get("Lang2", "")
        text1 = self.last_data.get("Text1", "")
        text2 = self.last_data.get("Text2", "")

        line = f"{timestamp},{lang1}:{text1},{lang2}:{text2}"

        log_path = os.path.join(LOG_DIR, "message.log")
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
            logging.info(f"Message logged: {line}")
        except Exception as e:
            logging.error(f"Log write error: {e}")

        self.current_id = None
        self.last_data = None
        self.last_update_time = None
message_buffer = MessageBuffer(
    stable_sec=config.get("PROCESS_STABLE_SEC", 10),
    flush_interval=config.get("FLUSH_INTERVAL_SEC", 5)
)


# ==============================
# WebSocket 受信 + 再接続
# ==============================
WS_RECONNECT_DELAY_SEC = config.get("WS_RECONNECT_DELAY_SEC", 5)
WS_MAX_RECONNECT_SEC = config.get("WS_MAX_RECONNECT_SEC", 60)

async def websocket_loop(url: str):
    notify("起動", "アプリケーションを起動しました")
    global ws

    # URL からポート番号を抽出
    try:
        port = url.split(":")[-1]
    except:
        port = "?"

    update_tray_status(f"接続待機中 (Port: {port})")

    start_retry_time = None

    while True:
        try:
            logging.info(f"Connecting WebSocket: {url}")
            update_tray_status(f"接続中… (Port: {port})")

            async with websockets.connect(url) as ws_local:
                ws = ws_local
                notify("データ受信準備完了", "WebSocket 接続が確立しました")
                logging.info("WebSocket connected")

                update_tray_status(f"受信中 (Port: {port})")

                start_retry_time = None

                async for message in ws_local:
                    if logging.getLogger().isEnabledFor(logging.DEBUG):
                        logging.debug(f"[WS RECV RAW] {message}")

                    await message_buffer.add_message(message)

        except Exception as e:
            logging.error(f"WebSocket error/disconnected: {e}")
            update_tray_status(f"再接続中… (Port: {port})")

            if start_retry_time is None:
                start_retry_time = time.time()

            elapsed = time.time() - start_retry_time
            if elapsed >= WS_MAX_RECONNECT_SEC:
                update_tray_status("切断")

            await asyncio.sleep(WS_RECONNECT_DELAY_SEC)

async def close_websocket():
    global ws
    try:
        if ws:
            await ws.close()
            logging.info("WebSocket closed.")
    except Exception as e:
        logging.error(f"Error closing WebSocket: {e}")


# ==============================
# タスクトレイ
# ==============================
def create_icon_image():
    img = Image.new("RGB", (16, 16), "blue")
    d = ImageDraw.Draw(img)
    d.rectangle([4, 4, 12, 12], fill="white")
    return img

def on_exit(icon, item):

    # WebSocket を閉じる（非同期タスクとして実行）
    loop = asyncio.get_event_loop()
    loop.create_task(safe_exit("user exit", 0))
    icon.stop()

def run_tray():
    global tray_icon
    tray_icon = pystray.Icon(
        "YukarinetteLogger",
        create_icon_image(),
        "接続待機中",
        menu=pystray.Menu(
            pystray.MenuItem("Exit（終了）", on_exit)
        ),
    )
    tray_icon.run()

def update_tray_status(text):
    global tray_icon
    if tray_icon:
        tray_icon.title = f"{APP_NAME}\n{text}"


# ==============================
# プロセス監視
# ==============================
async def process_monitor_loop():
    target = config.get("TARGET_PROCESS", "YNC_Neo.exe").lower()

    while True:
        found = False
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == target:
                    found = True
                    break
            except psutil.NoSuchProcess:
                pass

        if not found:
            await safe_exit("target process not found", 1)

        await asyncio.sleep(10)


# ==============================
# 終了処理
# ==============================
async def safe_exit(reason: str, code: int = 0):
    logging.info(f"Exit reason: {reason}")

    # WebSocket を閉じる
    try:
        await close_websocket()
    except:
        pass

    # メッセージバッファ flush
    async with message_buffer.lock:
        await message_buffer._flush_locked()

    notify("終了", f"アプリケーションを終了します ({reason})")

    os._exit(code)


# ==============================
# メイン
# ==============================
async def main_async():
    value = read_registry_value()

    if value is None:
        notify("終了", "レジストリから WebSocket の値を取得できませんでした")
        return

    if isinstance(value, int):
        ws_url = f"ws://127.0.0.1:{value}"
    else:
        ws_url = value

    flush_task = asyncio.create_task(message_buffer.periodic_flush())
    ws_task = asyncio.create_task(websocket_loop(ws_url))
    proc_task = asyncio.create_task(process_monitor_loop())  # ← 追加

    try:
        await ws_task
    finally:
        flush_task.cancel()
        proc_task.cancel()
        try:
            await flush_task
            await proc_task
        except asyncio.CancelledError:
            pass

def main():
    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()

    asyncio.run(main_async())


if __name__ == "__main__":
    main()
