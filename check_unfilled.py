#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
客户经理日志未填提醒 - 云端脚本 (设计运行于 GitHub Actions)

逻辑：
  1. 用飞书自建应用的 app_id/app_secret 换取 tenant_access_token
  2. 读取「客户经理名单」多维表格 -> 应填日志的客户经理名单(约31人)
  3. 读取「客户经理日志统计」多维表格，筛选当天(填写日期==今天) -> 已填集合
  4. 差集 = 今天未填的人
  5. 用同一自建应用，通过「发送消息」接口给你发一条飞书私信(单聊)

无需建群、无需群机器人 webhook。只要在飞书自建应用里打开「发送消息」权限，
并把应用加为两张表的协作者即可。

环境变量(强烈建议放 GitHub Secrets，不要写死在仓库里)：
  FEISHU_APP_ID          飞书自建应用 app_id
  FEISHU_APP_SECRET      飞书自建应用 app_secret
  FEISHU_USER_OPEN_ID    接收提醒的飞书用户 open_id (你的 open_id；
                        也可用 FEISHU_USER_ID + FEISHU_RECEIVE_ID_TYPE=user_id 指定工号)

说明：脚本不依赖本地 lark-cli，全部走飞书开放 API，适合在无头云端(CI)运行。
"""
import os
import sys
import re
import time
import json
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


_DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def _date_only(v):
    """从飞书日期字段值中提取 YYYY-MM-DD。

    飞书日期字段经 API 返回可能是 '2026-07-28 00:00:00'（带时分秒）、
    '2026-07-28'（纯日期）或 {'date': '2026-07-28'}（带结构），此处统一
    只取日期部分，避免与 today='2026-07-28' 做严格相等时因 ' 00:00:00'
    后缀而匹配失败。
    """
    s = _norm(v)
    m = _DATE_RE.search(s)
    return m.group(1) if m else s.strip()


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
    """按日期范围服务端筛选（today 00:00:00 <= 填写日期 < 次日 00:00:00）。

    飞书「填写日期」为日期时间型字段，返回带 ' 00:00:00' 后缀；用区间
    筛选（>= 当日0点 且 < 次日0点）可稳定命中当天全部记录，无需全量拉取。
    返回结果再用 _date_only 本地复核一次，双保险。
    """
    url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/search"
    y, m, d = (int(x) for x in date_str.split("-"))
    nxt = datetime.date(y, m, d) + datetime.timedelta(days=1)
    start = f"{date_str} 00:00:00"
    end = nxt.strftime("%Y-%m-%d") + " 00:00:00"
    out, page_token = [], ""
    while True:
        body = {
            "filter": {
                "conjunction": "and",
                "conditions": [
                    {"field_name": "填写日期", "operator": "isGreater", "value": [start]},
                    {"field_name": "填写日期", "operator": "isLess", "value": [end]},
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
            if _date_only(item.get("fields", {}).get("填写日期")) == date_str:
                out.append(item)
        if not data.get("data", {}).get("has_more") or not data.get("data", {}).get("page_token"):
            break
        page_token = data["data"]["page_token"]
    return out


def get_filled_names_today(token, today):
    """当天已填姓名集合。

    优先服务端按日期范围筛选（只取当天，省去全量拉取）；
    若服务端筛选抛异常，则回退全量拉取 + 本地 _date_only 过滤兜底。
    """
    try:
        items = search_records_by_date(LOG_APP, LOG_TABLE, token, today)
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 服务端日期筛选失败，回退全量拉取本地过滤: {e}")
        items = list_all_records(LOG_APP, LOG_TABLE, token)
    names = set()
    for item in items:
        f = item.get("fields", {})
        if _date_only(f.get("填写日期")) == today:
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


def send_app_message(token, receive_id, receive_id_type, text):
    """用自建应用给指定用户发一条飞书私信(单聊消息)。无需建群。"""
    url = f"{FEISHU_BASE}/im/v1/messages"
    params = {"receive_id_type": receive_id_type}
    body = {
        "receive_id": receive_id,
        "msg_type": "text",
        "content": json.dumps({"text": text}, ensure_ascii=False),
    }
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                      params=params, json=body, timeout=15)
    return r.json()


def main():
    app_id = os.environ.get("FEISHU_APP_ID")
    app_secret = os.environ.get("FEISHU_APP_SECRET")
    user_open_id = os.environ.get("FEISHU_USER_OPEN_ID")
    user_id = os.environ.get("FEISHU_USER_ID")

    # 接收人：优先用 open_id；若提供工号则回退到 user_id
    if user_open_id:
        receive_id, receive_id_type = user_open_id, "open_id"
    elif user_id:
        receive_id, receive_id_type = user_id, "user_id"
    else:
        receive_id, receive_id_type = None, None

    missing = [k for k, v in (("FEISHU_APP_ID", app_id), ("FEISHU_APP_SECRET", app_secret),
                              ("接收人(FEISHU_USER_OPEN_ID 或 FEISHU_USER_ID)", receive_id)) if not v]
    if missing:
        raise SystemExit(f"缺少环境变量: {', '.join(missing)}。请在 GitHub Secrets 中配置。")

    now = datetime.datetime.now(BJT)
    today = now.strftime("%Y-%m-%d")
    today_label = now.strftime("%Y年%m月%d日")

    token = get_tenant_token(app_id, app_secret)

    # 1) 应填名单（结合名单表「备注」：备注含"休假"当天免提醒）
    roster = list_all_records(ROSTER_APP, ROSTER_TABLE, token)
    names_raw, notes = [], {}
    for it in roster:
        f = it.get("fields", {})
        name = extract_name(f)
        if not name:
            continue
        names_raw.append(name)
        notes[name] = _norm(f.get("备注"))
    # 备注含"休假"视为当天不在岗，免提醒
    on_leave = {n for n in names_raw if "休假" in (notes.get(n) or "")}
    expected = sorted(set(names_raw) - on_leave - {""})
    print(f"[info] 名单 {len(names_raw)} 条，应提醒 {len(expected)} 人，休假免提醒 {len(on_leave)} 人：{sorted(on_leave)}")

    # 2) 已填（当天）
    filled = get_filled_names_today(token, today)

    # 3) 差集
    unfilled = [n for n in expected if n not in filled]

    if unfilled:
        lines = "\n".join(f"• {n}" for n in unfilled)
        msg = (f"【客户经理日志未填提醒】{today_label} 有 {len(unfilled)} 人未填\n"
               f"{lines}")
        if on_leave:
            msg += f"\n\n（备注含“休假”免提醒 {len(on_leave)} 人：{', '.join(sorted(on_leave))}）"
        msg += "\n\n请尽快在飞书补填今日日志。"
    else:
        extra = f"（备注含“休假”免提醒 {len(on_leave)} 人）" if on_leave else ""
        msg = f"✅ {today_label} 客户经理日志全员已填（应填 {len(expected)} 人）{extra}。"

    print(msg)  # 同时写进 Actions 日志，方便排查
    resp = send_app_message(token, receive_id, receive_id_type, msg)
    print("发送返回:", resp)


if __name__ == "__main__":
    main()
