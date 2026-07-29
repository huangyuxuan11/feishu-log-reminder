# -*- coding: utf-8 -*-
"""测试用：向本人飞书发送一条「定时触发器已生效」私信。"""
import os, json, datetime, requests

APP_ID = os.environ["FEISHU_APP_ID"]
APP_SECRET = os.environ["FEISHU_APP_SECRET"]
OPEN_ID = os.environ["FEISHU_USER_OPEN_ID"]


def get_token():
    r = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET},
        timeout=15,
    )
    r.raise_for_status()
    return r.json()["tenant_access_token"]


def send_text(text):
    token = get_token()
    resp = requests.post(
        "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id",
        headers={"Authorization": f"Bearer {token}"},
        json={
            "receive_id": OPEN_ID,
            "msg_type": "text",
            "content": json.dumps({"text": text}),
        },
        timeout=15,
    )
    print("feishu resp:", resp.status_code, resp.text)
    resp.raise_for_status()


if __name__ == "__main__":
    now_bjt = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))) \
        .strftime("%Y-%m-%d %H:%M:%S")
    send_text(
        f"✅ 定时触发测试\n触发时间：{now_bjt}（北京时间）\n"
        f"GitHub Actions 定时器已生效，飞书私信通道正常。"
    )
