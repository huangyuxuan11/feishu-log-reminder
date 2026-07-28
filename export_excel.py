#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
客户经理日志 - 每日Excel导出 (设计运行于 GitHub Actions，云端执行)

逻辑：
  1. 用飞书自建应用的 app_id/app_secret 换取 tenant_access_token
  2. 读取「客户经理日志统计」多维表格，筛选当天(填写日期==今天)的记录
  3. 把当天所有填报记录导出为 Excel（列：填写日期/客户经理姓名/工作时间1/工作内容1/工作时间2/工作内容2…，时段成对交叉）
  4. 通过飞书自建应用，把 Excel 作为文件私信发给你（亦非fan）

环境变量(放 GitHub Secrets)：
  FEISHU_APP_ID          飞书自建应用 app_id
  FEISHU_APP_SECRET      飞书自建应用 app_secret
  FEISHU_USER_OPEN_ID    接收人飞书 open_id（亦非fan）
  FEISHU_USER_ID         可选，留空即可

注意：发文件需要应用具备 im:resource 权限（文件上传）+ im:message + im:message:send_as_bot。
"""

import os
import io
import re
import json
import time
import datetime
import requests

# ---------- 配置：客户经理日志统计 多维表格 ----------
LOG_APP = "OteIbzKYha9jKKsprhwcKovNnCh"
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
    url = f"{FEISHU_BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    out, page_token = [], ""
    while True:
        params = {"page_size": page_size}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=20)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"读取记录失败: {data}")
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
                    {"field_name": "填写日期", "operator": ">=", "value": [start]},
                    {"field_name": "填写日期", "operator": "<", "value": [end]},
                ],
            },
            "page_size": page_size,
        }
        if page_token:
            body["page_token"] = page_token
        r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=body, timeout=20)
        data = r.json()
        if data.get("code") != 0:
            raise RuntimeError(f"筛选记录失败: {data}")
        for item in data.get("data", {}).get("items", []):
            if _date_only(item.get("fields", {}).get("填写日期")) == date_str:
                out.append(item)
        if not data.get("data", {}).get("has_more") or not data.get("data", {}).get("page_token"):
            break
        page_token = data["data"]["page_token"]
    return out


def get_today_records(token, today):
    """当天填报记录列表。

    优先服务端按日期范围筛选（只取当天，省去全量拉取）；
    若服务端筛选抛异常，则回退全量拉取 + 本地 _date_only 过滤兜底。
    """
    try:
        items = search_records_by_date(LOG_APP, LOG_TABLE, token, today)
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 服务端日期筛选失败，回退全量拉取本地过滤: {e}")
        items = list_all_records(LOG_APP, LOG_TABLE, token)
    return [it for it in items if _date_only(it.get("fields", {}).get("填写日期")) == today]


def build_excel(records, today_label):
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    # 找出最大的时段序号 N（字段形如 工作时间N / 工作内容N）
    max_n = 0
    for it in records:
        for k in it.get("fields", {}):
            m = re.match(r"^工作时间(\d+)$", k)
            if m:
                max_n = max(max_n, int(m.group(1)))
                continue
            m = re.match(r"^工作内容(\d+)$", k)
            if m:
                max_n = max(max_n, int(m.group(1)))

    # 多列、成对交叉：填写日期/客户经理姓名/工作时间1/工作内容1/工作时间2/工作内容2/...
    headers = ["填写日期", "客户经理姓名"]
    for n in range(1, max_n + 1):
        headers.append(f"工作时间{n}")
        headers.append(f"工作内容{n}")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "客户经理日志"

    # 表头样式
    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    # 数据行
    for r, it in enumerate(records, 2):
        f = it.get("fields", {})
        row_vals = [_norm(f.get("填写日期")), _norm(f.get("客户经理姓名"))]
        for n in range(1, max_n + 1):
            row_vals.append(_norm(f.get(f"工作时间{n}")))
            row_vals.append(_norm(f.get(f"工作内容{n}")))
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border

    # 列宽
    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 12
    for n in range(1, max_n + 1):
        tcol = 3 + (n - 1) * 2
        ccol = tcol + 1
        ws.column_dimensions[openpyxl.utils.get_column_letter(tcol)].width = 16
        ws.column_dimensions[openpyxl.utils.get_column_letter(ccol)].width = 40

    ws.freeze_panes = "C2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def upload_file(token, file_bytes, file_name):
    url = f"{FEISHU_BASE}/im/v1/files"
    files = {
        "file": (file_name, file_bytes,
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    }
    data = {"file_name": file_name, "file_type": "xlsx"}
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"},
                      files=files, data=data, timeout=30)
    resp = r.json()
    if resp.get("code") != 0:
        raise RuntimeError(f"文件上传失败: {resp}")
    return resp["data"]["file_key"]


def send_file_message(token, open_id, file_key):
    url = f"{FEISHU_BASE}/im/v1/messages?receive_id_type=open_id"
    body = {
        "receive_id": open_id,
        "msg_type": "file",
        "content": json.dumps({"file_key": file_key}),
    }
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=body, timeout=20)
    return r.json()


def send_text_message(token, open_id, text):
    url = f"{FEISHU_BASE}/im/v1/messages?receive_id_type=open_id"
    body = {
        "receive_id": open_id,
        "msg_type": "text",
        "content": json.dumps({"text": text}),
    }
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=body, timeout=20)
    return r.json()


def main():
    app_id = os.environ.get("FEISHU_APP_ID")
    app_secret = os.environ.get("FEISHU_APP_SECRET")
    open_id = os.environ.get("FEISHU_USER_OPEN_ID")
    if not app_id or not app_secret or not open_id:
        raise SystemExit("缺少环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_USER_OPEN_ID")

    now = datetime.datetime.now(BJT)
    today = now.strftime("%Y-%m-%d")
    today_label = now.strftime("%Y年%m月%d日")

    token = get_tenant_token(app_id, app_secret)
    records = get_today_records(token, today)
    print(f"[info] 当天({today})填报记录数: {len(records)}")

    if not records:
        msg = f"📭 {today_label} 暂无客户经理填写日志，未导出Excel。"
        print(msg)
        resp = send_text_message(token, open_id, msg)
        print("text 返回:", resp)
        return

    buf = build_excel(records, today_label)
    file_name = f"客户经理日志_{today}.xlsx"
    file_key = upload_file(token, buf.getvalue(), file_name)
    print(f"[info] 文件已上传 file_key={file_key}")

    # 先发一句说明，再发文件
    intro = f"📊 {today_label} 客户经理日志导出（共 {len(records)} 条填报），见附件Excel。"
    send_text_message(token, open_id, intro)
    resp = send_file_message(token, open_id, file_key)
    print("file 返回:", resp)
    print(f"✅ 已发送 {file_name}")


if __name__ == "__main__":
    main()
