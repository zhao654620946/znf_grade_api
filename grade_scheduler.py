import os
import sys
import time
import subprocess
import requests
import json
from datetime import datetime
from cryptography.fernet import Fernet
import base64
import hashlib


PUSH_URL = os.environ.get("PUSH_URL")
raw_key = os.environ.get("ENCRYPT_KEY")
if not raw_key:
    raise RuntimeError("ENCRYPT_KEY not found")

key = base64.urlsafe_b64encode(
    hashlib.sha256(raw_key.encode()).digest()
)

cipher = Fernet(key)

def send_push_notification(url, data):
    """发送推送消息"""
    response = requests.post(url, json=data)
    if response.status_code == 200:
        print("推送结果:", response.json())
    else:
        print("请求失败，状态码:", response.status_code)

def run_example_script():
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(current_dir, "znf_grade_api.py")
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True
        )

        stdout = result.stdout.decode("utf-8", errors="replace")
        stderr = result.stderr.decode("utf-8", errors="replace")

        if result.returncode == 0:
            return stdout.strip()
        else:
            print("执行失败：", stderr)
            return None

    except Exception as e:
        print("运行错误：", e)
        return None

def main():
    try:
        with open("data.txt", "rb") as f:
            encrypted = f.read()
        last_result = cipher.decrypt(encrypted).decode()

    except FileNotFoundError:
        last_result = None

    print()
    print("当前时间:", datetime.now())
    print("开始运行 znf_grade_api.py ...")
    current_result = run_example_script()

    if current_result is not None:  # 确保运行结果有效
        if last_result is None:
            print("首次运行，结果为：")
            # print(datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"+current_result)
            payload = {
                "title": "首次查询",  # 消息标题
                "content": datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"+current_result,
            }
            send_push_notification(PUSH_URL, payload)
        elif current_result != last_result:
            print("运行结果发生变化，新的结果为：")
            # print(datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"+current_result)
            payload = {
                "title": "成绩更新",  # 消息标题
                "content": datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"+current_result,
            }
            send_push_notification(PUSH_URL, payload)
        else:
            print("运行结果相同，无变化。结果为：")
            # print(datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"+current_result)

        # 等待 5 分钟
        encrypted_data = cipher.encrypt(current_result.encode())
        with open("data.txt", "wb") as f:
            f.write(encrypted_data)

if __name__ == "__main__":
    main()
