#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
客户经理日志未填提醒 - 云端脚本 (设计运行于 GitHub Actions)

逻辑：
  1. 用飞书自建应用的 app_id/app_secret 换取 tenant_access_token
  2. 读取「客户经理名单」多维表格 -> 应填日志的客户经理名单(约31人)
  3. 读取「客户经理日志统计」多维表格，筛选当天(填写日期==今天) -> 已填集合
  4. 差集 = 今天未填的人
  5. 通过飞书群机器人 webhook(HMAC-SHA256 签名) 把结果推到手机

环境变量(强烈建议放 GitHub Secrets，不要写死在仓库里)：
  FEISHU_APP_ID          飞书自建应用 app_id
  FEISHU_APP_SECRET      飞书自建应用 app_secret
  FEISHU_WEBHOOK         飞书群机器人 webhook 地址
  FEISHU_WEBHOOK_SECRET  飞书群机器人签名密钥(建议开启，防伪造)

说明：脚本不依赖本地 lark-cli，全部走飞书开放 API，适合在无头云端(CI)运行。
"""
import os
import sys
import time
import json
import hmac
import hashlib
import base64
import datetime
import requests

# ---------- 配置：两张飞书多维表格(已与本地 Excel 核对同源) ----------
ROSTER_APP = "DiFmbGm87aaPkxsmz6xctzWTnbe"   # 客户经理名单
ROSTER_TABLE = "tblQ3UVHeHiS8nZL"
LOG_APP = "OteIbzKYha9jKKsprhwcKovNnCh"      # 客户经理日志统计
LOG_TABLE = "tblAQ9ZGEkAmBgTT"

FEISHU_BASE = "https://open.feishu.cn/open-apis"
BJT = datetime.timezone(datetime.timedelta(hours=8))


def get_tenant_token(app_id, app_secret):
    url = f"{FEISHU_BASE}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=15)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"获取 tenant_access_token 失败: {data}")
    return data["tenant_access_token"]


def _norm(v):
    """把飞书字段值归一化为字符串(兼容 字符串/数组/带结构 的日期)。"""
    if v is None:
        return ""
    if isinstance(v, list):
        return _norm(v[0]) if v else ""
    if isinstance(v, dict):
        for k in ("date", "start"):
            if k in v:
                return str(v[k])
        return str(v)
    return str(v)


def list_all_records(app_token, table_id, token, page_size=100):
    """翻页拉取某表全部记录。"""
    url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    out, page_token = [], ""
    while True:
        params = {"page_size": page_size}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=20)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"读取记录失败 ({app_token}/{table_id}): {data}")
        out.extend(data.get("data", {}).get("items", []))
        if not data.get("data", {}).get("has_more") or not data.get("data", {}).get("page_token"):
            break
        page_token = data["data"]["page_token"]
    return out


def search_records_by_date(app_token, table_id, token, date_str, page_size=100):
    """按填写日期服务端精确筛选。"""
    url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/search"
    out, page_token = [], ""
    while True:
        body = {
            "filter": {
                "conjunction": "and",
                "conditions": [
                    {"field_name": "填写日期", "operator": "is", "value": [date_str]}
                ],
            },
            "page_size": page_size,
        }
        if page_token:
            body["page_token"] = page_token
        r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=body, timeout=20)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"筛选记录失败 ({app_token}/{table_id}): {data}")
        for item in data.get("data", {}).get("items", []):
            if _norm(item.get("fields", {}).get("填写日期")) == date_str:
                out.append(item)
        if not data.get("data", {}).get("has_more") or not data.get("data", {}).get("page_token"):
            break
        page_token = data["data"]["page_token"]
    return out


def get_filled_names_today(token, today):
    """当天已填姓名集合；优先服务端筛选，失败则全量拉取本地过滤(更稳)。"""
    try:
        items = search_records_by_date(LOG_APP, LOG_TABLE, token, today)
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 服务端日期筛选失败，改为全量拉取后本地过滤: {e}")
        items = list_all_records(LOG_APP, LOG_TABLE, token)
    names = set()
    for item in items:
        f = item.get("fields", {})
        if _norm(f.get("填写日期")) == today:
            n = extract_name(f)
            if n:
                names.add(n)
    return names


def extract_name(fields):
    """兼容 名单表'客户经理名称' 与 日志表'客户经理姓名'。"""
    for key in ("客户经理姓名", "客户经理名称"):
        if fields.get(key):
            return _norm(fields[key]).strip()
    for v in fields.values():
        s = _norm(v).strip()
        if s:
            return s
    return ""


def send_webhook(webhook, secret, text):
    ts = str(int(time.time()))
    sign = ""
    if secret:
        string_to_sign = ts + "\n" + secret
        hmac_code = hmac.new(secret.encode("utf-8"), string_to_sign.encode("utf-8"),
                             digestmod=hashlib.sha256).digest()
        sign = base64.b64encode(hmac_code).decode("utf-8")
    payload = {"msg_type": "text", "content": {"text": text}, "timestamp": ts}
    if sign:
        payload["sign"] = sign
    r = requests.post(webhook, json=payload, timeout=15)
    return r.json()


def main():
    app_id = os.environ.get("FEISHU_APP_ID")
    app_secret = os.environ.get("FEISHU_APP_SECRET")
    webhook = os.environ.get("FEISHU_WEBHOOK")
    secret = os.environ.get("FEISHU_WEBHOOK_SECRET", "")

    missing = [k for k, v in (("FEISHU_APP_ID", app_id), ("FEISHU_APP_SECRET", app_secret),
                              ("FEISHU_WEBHOOK", webhook)) if not v]
    if missing:
        raise SystemExit(f"缺少环境变量: {', '.join(missing)}。请在 GitHub Secrets 中配置。")

    now = datetime.datetime.now(BJT)
    today = now.strftime("%Y-%m-%d")
    today_label = now.strftime("%Y年%m月%d日")

    token = get_tenant_token(app_id, app_secret)

    # 1) 应填名单
    roster = list_all_records(ROSTER_APP, ROSTER_TABLE, token)
    expected = sorted({extract_name(it.get("fields", {})) for it in roster} - {""})

    # 2) 已填（当天）
    filled = get_filled_names_today(token, today)

    # 3) 差集
    unfilled = [n for n in expected if n not in filled]

    if unfilled:
        lines = "\n".join(f"• {n}" for n in unfilled)
        msg = (f"【客户经理日志未填提醒】{today_label} 有 {len(unfilled)} 人未填\n"
               f"{lines}\n\n请尽快在飞书补填今日日志。")
    else:
        msg = f"✅ {today_label} 客户经理日志全员已填（应填 {len(expected)} 人）。"

    print(msg)  # 同时写进 Actions 日志，方便排查
    resp = send_webhook(webhook, secret, msg)
    print("webhook 返回:", resp)


if __name__ == "__main__":
    main()
