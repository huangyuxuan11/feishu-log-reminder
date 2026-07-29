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

    飞书日期时间型字段经开放 API 返回格式不固定：
      - 整数/浮点毫秒时间戳（如 1785168000000）→ 按 BJT 时区转 datetime
      - 字符串 '2026-07-28 00:00:00' / '2026-07-28' → 正则提取日期部分
      - dict {'date': ...} / list [...] → 取内部值
    统一返回 'YYYY-MM-DD'，用于与 today 字符串比较。
    """
    # dict / list 展开
    if isinstance(v, dict):
        for k in ("date", "start"):
            if k in v:
                v = v[k]
                break
    if isinstance(v, list):
        v = v[0] if v else ""
    # 整数/浮点 → 视为毫秒时间戳
    if isinstance(v, bool):
        pass
    elif isinstance(v, (int, float)):
        try:
            return datetime.datetime.fromtimestamp(v / 1000, tz=BJT).strftime("%Y-%m-%d")
        except (ValueError, OSError, OverflowError):
            return str(v)
    s = _norm(v) if not isinstance(v, str) else v
    if isinstance(s, str):
        s = s.strip()
        m = _DATE_RE.search(s)
        return m.group(1) if m else s
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


# ---------- 去重控制表（避免飞书自动化 与 GitHub 自带定时器 双触发导致重复发消息） ----------
CONTROL_APP = "H0qGbDCJiaMXQJss7xOcvTYEnsh"   # 日志提醒去重控制
CONTROL_TABLE = "tbleEVzbXaicZ19E"

# 北京时间各触发档位（与飞书自动化 / GitHub schedule 一致）
SLOTS = [(2000, 20, 0), (2030, 20, 30), (2045, 20, 45), (2130, 21, 30)]


def get_slot(now):
    """返回当前时间最接近的档位 key（±20 分钟内），否则 None（如手动测试）。"""
    cur = now.hour * 60 + now.minute
    best, best_diff = None, 999
    for key, h, m in SLOTS:
        diff = abs(cur - (h * 60 + m))
        if diff < best_diff:
            best, best_diff = key, diff
    return best if best_diff <= 20 else None


def _create_record(app_token, table_id, fields, token):
    url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json={"fields": fields}, timeout=20)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"创建记录失败: {data}")
    return data.get("data", {}).get("record")


def _update_record(app_token, table_id, record_id, fields, token):
    url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}"
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json={"fields": fields}, timeout=20)
    data = r.json()
    if data.get("code") != 0:
        raise RuntimeError(f"更新记录失败: {data}")
    return data


def _find_control_record(records, task):
    for it in records:
        raw = _norm(it.get("fields", {}).get("文本"))
        if raw.startswith(f"{task}|"):
            return it, raw
    return None, ""


def read_fired_slots(token, task, today):
    """读取该任务今天已发送的档位集合；异常时返回空集（fail-open，不阻断发送）。"""
    try:
        records = list_all_records(CONTROL_APP, CONTROL_TABLE, token)
        it, raw = _find_control_record(records, task)
        if it and raw:
            parts = raw.split("|")
            if len(parts) >= 3 and parts[1] == today:
                return set(p for p in parts[2].split(",") if p)
    except Exception as e:
        print(f"[dedup] 读取控制表失败(忽略，继续发送): {e}")
    return set()


def mark_slot_fired(token, task, today, slot):
    """标记该任务今天某档位已发送；异常时仅记录（fail-open）。"""
    try:
        records = list_all_records(CONTROL_APP, CONTROL_TABLE, token)
        it, raw = _find_control_record(records, task)
        if it:
            rec_id = it.get("record_id") or it.get("id")
            parts = raw.split("|")
            if len(parts) >= 3 and parts[1] == today:
                slots = set(p for p in parts[2].split(",") if p)
                slots.add(str(slot))
                _update_record(CONTROL_APP, CONTROL_TABLE, rec_id,
                               {"文本": f"{task}|{today}|{','.join(sorted(slots))}"}, token)
            else:
                _update_record(CONTROL_APP, CONTROL_TABLE, rec_id,
                               {"文本": f"{task}|{today}|{slot}"}, token)
        else:
            _create_record(CONTROL_APP, CONTROL_TABLE, {"文本": f"{task}|{today}|{slot}"}, token)
    except Exception as e:
        print(f"[dedup] 写入控制表失败(忽略): {e}")


def get_filled_names_today(token, today):
    """当天已填姓名集合。

    飞书日期时间型字段的服务端筛选(isGreater/isLess)对 value 格式要求苛刻
    且文档不明确，多次尝试均报 InvalidFilter。
    故采用稳定策略：全量拉取 + 本地 _date_only 过滤。
    """
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

    # 0) 去重：若本档位今日已由另一触发器（飞书自动化 / GitHub 自带定时器）发送过，跳过防重复
    slot = get_slot(now)
    if slot is not None:
        if str(slot) in read_fired_slots(token, "remind", today):
            print(f"[dedup] 档位 {slot} 今日已发送，跳过（防重复）")
            return

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
    if slot is not None:
        mark_slot_fired(token, "remind", today, slot)


if __name__ == "__main__":
    main()
