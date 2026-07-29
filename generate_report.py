#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
客户经理日志分析报告 - 云端自动生成 (设计运行于 GitHub Actions)

逻辑（与 未填提醒 / 每日导出 同一模式）：
  1. 用飞书自建应用的 app_id/app_secret 换取 tenant_access_token
  2. 读取「客户经理名单」多维表格 -> 应填名单 + 组长 -> 分局 映射
  3. 读取「客户经理日志统计」多维表格，筛选当天(填写日期==今天)的填报记录
  4. 自动统计：段数 / 分局归属 / 工作摘要(简化) / 提交率 / 拜访冠军 / 分局排行
     并基于关键词扫描自动抽取 签约落地 / 关键商机 / 产品维度 / 明日关注 候选条目
  5. 渲染成与本地美化版一致风格的 HTML，再用无头 Chromium 转 PDF
  6. 用同一自建应用，把 PDF 作为文件私信发给你（与导出 Excel 同一模式）

环境变量(放 GitHub Secrets，复用现有三个即可)：
  FEISHU_APP_ID          飞书自建应用 app_id
  FEISHU_APP_SECRET      飞书自建应用 app_secret
  FEISHU_USER_OPEN_ID    接收人飞书 open_id（亦非fan）
  FEISHU_USER_ID         可选
  REPORT_DATE            可选；覆盖报告日期(YYYY-MM-DD)，用于排错/补生成
  SELFTEST               可选；=1 时用内置假数据生成预览，不连飞书、不发消息

依赖：requests / openpyxl（已有）+ playwright（云端转 PDF 用）
"""
import os
import re
import io
import sys
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

# 组长 -> 分局/组（组织架构固定，Emoji 为分局配色圆标，color/dark 为标题渐变主色）
BRANCHES = [
    ("郑文博", "洪武路五老村分局", "🔴", "#e74c3c", "#c0392b", "#fff"),
    ("裴想",   "朝天宫双塘分局",   "🟠", "#f39c12", "#d68910", "#fff"),
    ("董兰凤", "大光瑞金月牙湖分局", "🟡", "#f1c40f", "#caa60c", "#3a2f00"),
    ("许仁珏", "红花中华门分局",   "🟢", "#27ae60", "#1e8449", "#fff"),
    ("王益华", "光华路分局",       "🔵", "#3498db", "#2471a3", "#fff"),
    ("刘芾源", "夫子庙秦虹分局",   "🟣", "#9b59b6", "#76448a", "#fff"),
    ("殷若婷", "行业组",           "⚪", "#95a5a6", "#707b7c", "#fff"),
]
LEADER_TO_BRANCH = {b[0]: b for b in BRANCHES}
DEFAULT_BRANCH = ("未归属", "其他 / 未归属", "❔", "#7f8c8d", "#616a6b", "#fff")

# 产品维度关键词
PRODUCTS = [
    ("一网通", ["一网通"]),
    ("专线/宽带", ["专线", "宽带"]),
    ("统付", ["统付"]),
    ("欠费催缴", ["欠费", "催缴"]),
    ("物联网卡", ["物联网"]),
    ("酒店摸排", ["酒店"]),
    ("明厨亮灶", ["明厨"]),
    ("5G融媒", ["5G", "融媒"]),
    ("V网", ["V网"]),
    ("SaaS", ["saas", "SaaS"]),
    ("云/AI", ["云", "AI", "ai"]),
]

# ---------- 去重控制表（与提醒/导出共用同一张表，任务名 report） ----------
CONTROL_APP = "H0qGbDCJiaMXQJss7xOcvTYEnsh"   # 日志提醒去重控制
CONTROL_TABLE = "tbleEVzbXaicZ19E"
SLOTS = [(2200, 22, 0)]   # 报告触发档位：北京时间 22:00


# ===================== 飞书基础工具（与现有脚本一致） =====================
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
    if isinstance(v, dict):
        for k in ("date", "start"):
            if k in v:
                v = v[k]
                break
    if isinstance(v, list):
        v = v[0] if v else ""
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


# ===================== 去重（与现有脚本一致） =====================
def get_slot(now):
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
    try:
        records = list_all_records(CONTROL_APP, CONTROL_TABLE, token)
        it, raw = _find_control_record(records, task)
        if it and raw:
            parts = raw.split("|")
            if len(parts) >= 3 and parts[1] == today:
                return set(p for p in parts[2].split(",") if p)
    except Exception as e:
        print(f"[dedup] 读取控制表失败(忽略，继续生成): {e}")
    return set()


def mark_slot_fired(token, task, today, slot):
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


# ===================== 数据解析工具 =====================
_TIME_ONLY_RE = re.compile(r"^[\s]*(\d{1,2}[:：]\d{2}|\d{3,4}[-~～]\d{3,4}|\d{1,2}\.\d{2}|\d{1,2}点到\d{1,2}点)[\s]*$")
_PURE_MEETING = {"晨会", "早会", "例会", "班会", "周会", "夕会", "午会"}


def extract_name(fields):
    for key in ("客户经理姓名", "客户经理名称"):
        if fields.get(key):
            return _norm(fields[key]).strip()
    for v in fields.values():
        s = _norm(v).strip()
        if s:
            return s
    return ""


def max_seg_index(fields):
    n = 0
    for k in fields:
        m = re.match(r"^工作时间(\d+)$", k)
        if m:
            n = max(n, int(m.group(1)))
        m = re.match(r"^工作内容(\d+)$", k)
        if m:
            n = max(n, int(m.group(1)))
    return n


def person_segments(fields):
    """段数 = 含有效工作内容的 时间/内容 配对数量。"""
    n = max_seg_index(fields)
    cnt = 0
    for i in range(1, n + 1):
        content = _norm(fields.get(f"工作内容{i}")).strip()
        if content:
            cnt += 1
    return cnt


def person_summaries(fields, max_items=5, trunc=18):
    """工作摘要：取每条工作内容，剥噪声，截断，最多前5条，超出显示「等N项」。"""
    n = max_seg_index(fields)
    items = []
    for i in range(1, n + 1):
        c = _norm(fields.get(f"工作内容{i}")).strip()
        if not c:
            continue
        if _TIME_ONLY_RE.match(c):          # 纯时间段残留
            continue
        if c in _PURE_MEETING:              # 纯会议占位词
            continue
        c = c[:trunc] + ("…" if len(c) > trunc else "")
        items.append(c)
    if len(items) > max_items:
        return items[:max_items] + [f"等{len(items) - max_items}项"]
    return items


def person_contents(fields):
    """原始工作内容列表（用于关键词扫描签约/商机/产品/明日）。"""
    n = max_seg_index(fields)
    out = []
    for i in range(1, n + 1):
        c = _norm(fields.get(f"工作内容{i}")).strip()
        if c:
            out.append(c)
    return out


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def get_weekday(d):
    return ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][d.weekday()]


# ===================== 统计主流程 =====================
def build_stats(today, roster_items, log_items):
    # 名单映射：name -> (组长, 备注)
    roster_map = {}
    for it in roster_items:
        f = it.get("fields", {})
        name = extract_name(f)
        if name:
            roster_map[name] = (_norm(f.get("客户经理组长")), _norm(f.get("备注")))

    # 当天填报
    today_records = [it for it in log_items if _date_only(it.get("fields", {}).get("填写日期")) == today]
    submit_map = {}   # name -> fields
    for it in today_records:
        name = extract_name(it.get("fields", {}))
        if name:
            submit_map[name] = it.get("fields", {})

    # 人员结构
    people = []  # dict per person
    for name, (leader, note) in roster_map.items():
        on_leave = "休假" in (note or "")
        branch = LEADER_TO_BRANCH.get(leader, DEFAULT_BRANCH)
        f = submit_map.get(name)
        if f:
            segs = person_segments(f)
            summ = person_summaries(f)
            raw = person_contents(f)
            status = "提交"
        else:
            segs, summ, raw = 0, [], []
            status = "休假" if on_leave else "未提交"
        people.append({
            "name": name, "leader": leader, "branch": branch,
            "on_leave": on_leave, "status": status,
            "segs": segs, "summ": summ, "raw": raw,
        })

    # 未归属（填报了但不在名单里）
    for name, f in submit_map.items():
        if name not in roster_map:
            people.append({
                "name": name, "leader": "", "branch": DEFAULT_BRANCH,
                "on_leave": False, "status": "提交",
                "segs": person_segments(f), "summ": person_summaries(f),
                "raw": person_contents(f),
            })

    total = len(roster_map)
    on_leave_names = [p["name"] for p in people if p["on_leave"]]
    expected = total - len(on_leave_names)
    submitted = [p for p in people if p["status"] == "提交"]
    submit_rate = (len(submitted) / expected * 100) if expected else 0

    # 拜访冠军 / 卷王 / 高效组
    sub_sorted = sorted(submitted, key=lambda p: p["segs"], reverse=True)
    champion = sub_sorted[0] if sub_sorted else None
    six_juan = [p for p in sub_sorted if p["segs"] == 6 and champion and p["name"] != champion["name"]]
    five_group = [p["name"] for p in sub_sorted if p["segs"] == 5]

    # 分局聚合
    branches = []
    for b in BRANCHES:
        members = [p for p in people if p["branch"][1] == b[1]]
        b_sub = [m for m in members if m["status"] == "提交"]
        b_total = sum(m["segs"] for m in b_sub)
        b_expected = [m for m in members if not m["on_leave"]]
        b_rate = (len(b_sub) / len(b_expected) * 100) if b_expected else 0
        branches.append({"info": b, "members": members, "submitted": b_sub,
                         "total_segs": b_total, "rate": b_rate, "expected": len(b_expected)})

    # 关键词扫描
    def scan(keywords):
        res = []
        for p in submitted:
            for line in p["raw"]:
                if any(k in line for k in keywords):
                    res.append({"name": p["name"], "branch": p["branch"][1], "line": line})
        return res

    signs = scan(["签约", "落地", "✅"])
    chances = scan(["商机", "意向", "🔥"])
    tomorrow = scan(["明日", "明天", "关注", "跟进"])

    # 产品维度
    products = []
    for pname, kws in PRODUCTS:
        owners = set()
        branches_cov = set()
        for p in submitted:
            if any(k in c for c in p["raw"] for k in kws):
                owners.add(p["name"])
                branches_cov.add(p["branch"][1])
        if owners:
            products.append({"name": pname, "branches": branches_cov, "count": len(owners),
                             "rep": ", ".join(sorted(owners)[:3])})

    total_segs = sum(p["segs"] for p in submitted)

    # 走访强度提示：仅统计应出勤（非休假）且当日已提交日志的客户经理
    #   1-2段 = 走访强度偏弱，建议加强；3段 = 基本达标，不批评；4段及以上不点名
    submitted_active = [p for p in submitted if not p["on_leave"]]
    visit_weak = [p for p in submitted_active if p["segs"] <= 2]
    visit_basic = [p for p in submitted_active if p["segs"] == 3]

    return {
        "today": today, "people": people, "total": total,
        "expected": expected, "submitted": submitted, "submit_rate": submit_rate,
        "on_leave": on_leave_names, "champion": champion, "six_juan": six_juan,
        "five_group": five_group, "branches": branches, "signs": signs,
        "chances": chances, "tomorrow": tomorrow, "products": products,
        "total_segs": total_segs, "visit_weak": visit_weak, "visit_basic": visit_basic,
    }


# ===================== HTML 渲染 =====================
CSS = """
* { box-sizing: border-box; }
body { font-family: "Microsoft YaHei","PingFang SC","Noto Sans CJK SC",sans-serif; color:#1a1a2e; margin:0; padding:24px 28px; background:#f5f7fa; }
h1 { font-size:22px; margin:0 0 2px; }
.sub { color:#888; font-size:13px; margin-bottom:18px; }
.section { background:#fff; border-radius:12px; padding:16px 18px; margin-bottom:16px; box-shadow:0 1px 3px rgba(0,0,0,.06); }
.section h2 { font-size:16px; margin:0 0 12px; padding-left:10px; border-left:4px solid #3498db; }
.ov { display:flex; flex-wrap:wrap; gap:12px; }
.card { flex:1 1 30%; min-width:180px; background:#fff; border:1px solid #e8ecf1; border-radius:10px; padding:14px 16px; border-left:4px solid #3498db; }
.ck { font-size:12px; color:#888; margin-bottom:4px; }
.cv { font-size:16px; font-weight:700; color:#1a1a2e; }
.branch-h { padding:10px 16px; font-size:15px; font-weight:700; color:#fff; border-radius:10px 10px 0 0; margin-top:6px; }
.branch-body { border:1px solid #eee; border-top:none; border-radius:0 0 10px 10px; padding:8px 12px 12px; }
table { width:100%; border-collapse:collapse; font-size:13px; margin-top:6px; }
th,td { border:1px solid #eee; padding:6px 8px; text-align:left; vertical-align:top; }
th { background:#f7f9fc; font-weight:700; }
.badge { display:inline-block; background:#3498db; color:#fff; border-radius:12px; padding:2px 10px; font-size:12px; font-weight:700; }
.hl { margin:8px 4px 4px; padding:8px 14px; background:#fffbf0; border-left:4px solid #f1c40f; border-radius:0 6px 6px 0; font-size:12px; color:#666; }
.sign-tbl thead th { background:#eafaf1!important; color:#27ae60!important; }
.prod-tbl thead th { background:#ebf5fb!important; color:#2980b9!important; }
.rank-tbl thead th { background:#f3e5f5!important; color:#8e44ad!important; }
.tm { list-style:none; padding:0; margin:0; }
.tm li { display:flex; align-items:flex-start; padding:6px 0; border-bottom:1px solid #f0f0f0; font-size:13px; }
.num { display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:50%; background:#e74c3c; color:#fff; font-size:11px; font-weight:700; margin-right:8px; flex:0 0 auto; }
.note { font-size:12px; color:#999; margin-top:6px; }
.notes li { font-size:13px; color:#555; margin:3px 0; }
.visit-block { margin-top:10px; border:1px solid #eef1f5; border-radius:10px; padding:10px 14px; }
.visit-sub { font-size:13px; font-weight:700; margin-bottom:6px; }
.visit-sub.weak { color:#c0392b; }
.visit-sub.ok { color:#27ae60; }
.vnum { display:inline-flex; align-items:center; justify-content:center; width:20px; height:20px; border-radius:50%; background:#e74c3c; color:#fff; font-size:11px; font-weight:700; margin-right:8px; flex:0 0 auto; }
.vnum.ok { background:#27ae60; }
"""


def generate_html(stats, today_label):
    s = stats
    champ = s["champion"]
    champ_txt = f"{champ['name']}（{champ['segs']}段）" if champ else "—"
    six_txt = "、".join(p["name"] for p in s["six_juan"]) or "无"
    five_txt = "、".join(s["five_group"]) or "无"

    # 全景概览卡片
    ov_cards = [
        ("团队总数", f"{s['total']}人（6分局+行业组）"),
        ("应出勤 / 实际提交", f"{s['expected']} / {len(s['submitted'])}（{s['submit_rate']:.0f}%）"),
        ("总段数", f"{s['total_segs']}段"),
        ("拜访冠军🥇", champ_txt),
        ("6段卷王", six_txt),
        ("5段高效组", five_txt),
    ]
    ov_html = '<div class="ov">' + "".join(
        f'<div class="card"><div class="ck">{esc(k)}</div><div class="cv">{esc(v)}</div></div>'
        for k, v in ov_cards) + '</div>'

    # 分局深度分析
    branch_html = ""
    for b in s["branches"]:
        info = b["info"]
        emoji, name, color, dark, txt = info[2], info[1], info[3], info[4], info[5]
        rate = f"{b['rate']:.0f}%"
        header = (f'<div class="branch-h" style="background:linear-gradient(135deg,{color},{dark});color:{txt};">'
                  f'{emoji} {esc(name)}（{b["expected"]}人应出勤·{len(b["submitted"])}人提交·{rate}）</div>')
        rows = ""
        for m in sorted([x for x in b["members"] if not x["on_leave"]],
                        key=lambda x: (x["status"] != "提交", -x["segs"])):
            if m["status"] == "提交":
                summ = " → ".join(m["summ"]) if m["summ"] else "（无明细）"
                segs_cell = f'<span class="badge">{m["segs"]}段</span>'
            else:
                summ = "未提交"
                segs_cell = '<span class="badge" style="background:#e74c3c;">未提交</span>'
            rows += (f'<tr><td style="width:90px;font-weight:600;">{esc(m["name"])}</td>'
                     f'<td style="width:64px;">{segs_cell}</td>'
                     f'<td>{esc(summ)}</td></tr>')
        if not rows:
            rows = '<tr><td colspan="3" style="color:#999;">（当日该分局无应出勤人员）</td></tr>'
        table = (f'<table><thead><tr><th>成员</th><th>段数</th><th>工作摘要</th></tr></thead>'
                 f'<tbody>{rows}</tbody></table>')
        # 分局亮点（数据派生）
        if b["submitted"]:
            top = max(b["submitted"], key=lambda x: x["segs"])
            hl = f"本分局提交 {len(b['submitted'])}/{b['expected']} 人，共 {b['total_segs']} 段；段数最高 {top['name']}（{top['segs']}段）。"
        else:
            hl = "本分局当日暂无提交。"
        branch_html += (f'<div style="margin-bottom:12px;">{header}'
                        f'<div class="branch-body">{table}'
                        f'<div class="hl">分局亮点：{esc(hl)}</div></div></div>')

    # 签约落地
    if s["signs"]:
        sign_rows = "".join(
            f'<tr><td>{esc(r["name"])}</td><td>{esc(r["branch"])}</td><td>{esc(r["line"])}</td></tr>'
            for r in s["signs"])
        sign_html = (f'<table class="sign-tbl"><thead><tr><th>客户经理</th><th>分局</th><th>签约/落地事项</th></tr></thead>'
                     f'<tbody>{sign_rows}</tbody></table>')
    else:
        sign_html = '<p class="note">当日日志未提取到签约/落地相关条目（如需，可人工甄选后补充）。</p>'

    # 关键商机
    if s["chances"]:
        ch_rows = "".join(
            f'<tr><td>{esc(r["name"])}</td><td>{esc(r["branch"])}</td><td>{esc(r["line"])}</td></tr>'
            for r in s["chances"])
        ch_html = (f'<table><thead><tr><th>客户经理</th><th>分局</th><th>商机/意向</th></tr></thead>'
                   f'<tbody>{ch_rows}</tbody></table>')
    else:
        ch_html = '<p class="note">当日日志未提取到商机/意向相关条目。</p>'

    # 产品维度
    if s["products"]:
        prod_rows = "".join(
            f'<tr><td>{esc(p["name"])}</td><td>{esc("、".join(p["branches"]))}</td>'
            f'<td>{p["count"]}人</td><td>{esc(p["rep"])}</td></tr>'
            for p in s["products"])
        prod_html = (f'<table class="prod-tbl"><thead><tr><th>产品线</th><th>覆盖分局</th><th>覆盖人数</th><th>代表</th></tr></thead>'
                     f'<tbody>{prod_rows}</tbody></table>')
    else:
        prod_html = '<p class="note">当日日志未提取到产品相关关键词。</p>'

    # 明日关注
    if s["tomorrow"]:
        tm_items = "".join(
            f'<li><span class="num">•</span><div><b>{esc(r["name"])}（{esc(r["branch"])}）</b>：{esc(r["line"])}</div></li>'
            for r in s["tomorrow"])
        tm_html = f'<ul class="tm">{tm_items}</ul>'
    else:
        tm_html = '<p class="note">当日日志未提取到明日跟进事项（可人工补充）。</p>'

    # 分局排行
    ranked = sorted(s["branches"], key=lambda b: b["total_segs"], reverse=True)
    medals = ["🥇", "🥈", "🥉"]
    rank_rows = ""
    for i, b in enumerate(ranked):
        emoji = b["info"][2]
        rank = medals[i] if i < 3 else str(i + 1)
        rank_rows += (f'<tr><td>{rank}</td><td>{emoji} {esc(b["info"][1])}</td>'
                      f'<td>{b["rate"]:.0f}%</td><td>{b["total_segs"]}段</td>'
                      f'<td>{len([x for x in b["submitted"] if any(k in c for c in x["raw"] for k in ["签约","落地"])])}单</td>'
                      f'<td>{esc(b["info"][1])}当日共{b["total_segs"]}段</td></tr>')
    rank_html = (f'<table class="rank-tbl"><thead><tr><th>排名</th><th>分局</th><th>提交率</th><th>总段数</th><th>签约</th><th>关键商机</th></tr></thead>'
                 f'<tbody>{rank_rows}</tbody></table>')

    # 数据说明
    notes = [
        f"统计周期：{esc(today_label)}（北京时间）",
        f"团队总数：{s['total']}人（6分局+行业组）",
        f"应出勤：{s['expected']}人（备注含“休假”免统计 {len(s['on_leave'])} 人：{esc('、'.join(s['on_leave']) or '无')}）",
        f"实际提交：{len(s['submitted'])}人，提交率 {s['submit_rate']:.0f}%",
        f"总段数：{s['total_segs']}段（系统按“时间+内容”配对自动统计）",
        "备注：签约/商机/明日关注/产品维度由脚本按关键词自动抽取，供参考；如需正式口径请以人工甄选版为准。",
        "数据来源：飞书多维表格「客户经理日志统计」（与本地Excel同源）。",
    ]
    notes_html = '<ul class="notes">' + "".join(f"<li>{n}</li>" for n in notes) + "</ul>"

    # 九、走访强度提示（固定模块：委婉点名工作量不饱和者，中性表述）
    weak, basic = s["visit_weak"], s["visit_basic"]
    if weak or basic:
        weak_items = "".join(
            f'<li><span class="vnum">!</span><div><b>{esc(p["name"])}</b>（{p["segs"]}段）：'
            f'走访密度偏低，建议在系统中分时段记录、细化记录颗粒度，并加大摸排走访频次。</div></li>'
            for p in weak)
        basic_items = "".join(
            f'<li><span class="vnum ok">✓</span><div>{esc(p["name"])}（{p["segs"]}段）：基本达标，维持现有走访节奏即可。</div></li>'
            for p in basic)
        visit_html = (
            f'<p class="note">说明：本模块仅统计应出勤（备注不含“休假”）且当日已提交日志的客户经理；'
            f'休假人员免统计、未提交人员在分局分析中已单列，此处不重复通报。</p>'
            f'<div class="visit-block"><div class="visit-sub weak">走访强度偏弱（1–2段，建议加强）</div>'
            f'<ul class="tm">{weak_items}</ul></div>'
            f'<div class="visit-block"><div class="visit-sub ok">基本达标（3段）</div>'
            f'<ul class="tm">{basic_items}</ul></div>')
    else:
        visit_html = '<p class="note">当日提交人员走访强度整体良好，无偏弱情况，不点名。</p>'

    html = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<title>客户经理日志分析报告 {esc(today_label)}</title><style>{CSS}</style></head>
<body>
<h1>客户经理日志分析报告 | {esc(today_label)}</h1>
<div class="sub">数据来源：客户经理日志统计（飞书多维表格）</div>

<div class="section"><h2>一、全景概览</h2>{ov_html}</div>
<div class="section"><h2>二、分局深度分析</h2>{branch_html}</div>
<div class="section"><h2>三、签约落地 ✅</h2>{sign_html}</div>
<div class="section"><h2>四、关键商机追踪 🔥</h2>{ch_html}</div>
<div class="section"><h2>五、产品维度分析</h2>{prod_html}</div>
<div class="section"><h2>六、明日关注</h2>{tm_html}</div>
<div class="section"><h2>七、分局活跃度排行</h2>{rank_html}</div>
<div class="section"><h2>八、数据说明</h2>{notes_html}</div>
<div class="section"><h2>九、走访强度提示</h2>{visit_html}</div>
</body></html>"""
    return html


# ===================== PDF（无头 Chromium，云端/本地通用） =====================
def html_to_pdf(html_path, pdf_path):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox", "--disable-gpu"])
        page = p.chromium.new_page()
        page.goto("file://" + html_path)
        page.pdf(path=pdf_path, format="A4", print_background=True,
                 margin={"top": "12mm", "bottom": "12mm", "left": "10mm", "right": "10mm"})
        browser.close()


# ===================== 飞书发送（与导出一致） =====================
def upload_file(token, file_bytes, file_name, file_type, mime):
    url = f"{FEISHU_BASE}/im/v1/files"
    files = {"file": (file_name, file_bytes, mime)}
    data = {"file_name": file_name, "file_type": file_type}
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, files=files, data=data, timeout=30)
    resp = r.json()
    if resp.get("code") != 0:
        raise RuntimeError(f"文件上传失败: {resp}")
    return resp["data"]["file_key"]


def send_file_message(token, open_id, file_key):
    url = f"{FEISHU_BASE}/im/v1/messages?receive_id_type=open_id"
    body = {"receive_id": open_id, "msg_type": "file", "content": json.dumps({"file_key": file_key})}
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=body, timeout=20)
    return r.json()


def send_text_message(token, open_id, text):
    url = f"{FEISHU_BASE}/im/v1/messages?receive_id_type=open_id"
    body = {"receive_id": open_id, "msg_type": "text", "content": json.dumps({"text": text})}
    r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json=body, timeout=20)
    return r.json()


# ===================== 自测（假数据，不连飞书） =====================
def selftest():
    today = "2026-07-29"
    roster = [
        {"fields": {"客户经理名称": "张三", "客户经理组长": "郑文博", "备注": ""}},
        {"fields": {"客户经理名称": "李四", "客户经理组长": "郑文博", "备注": ""}},
        {"fields": {"客户经理名称": "王五", "客户经理组长": "裴想", "备注": "休假"}},
        {"fields": {"客户经理名称": "赵六", "客户经理组长": "董兰凤", "备注": ""}},
        {"fields": {"客户经理名称": "孙七", "客户经理组长": "许仁珏", "备注": ""}},
    ]
    logs = [
        {"fields": {"填写日期": "2026-07-29", "客户经理姓名": "张三",
                    "工作时间1": "9:00", "工作内容1": "拜访南京雅高资产客户，推介一网通",
                    "工作时间2": "14:00", "工作内容2": "签约宽带专线1条✅"}},
        {"fields": {"填写日期": "2026-07-29", "客户经理姓名": "李四",
                    "工作时间1": "10:00", "工作内容1": "晨会",
                    "工作时间2": "11:00", "工作内容2": "摸排红花片区酒店，商机意向2家🔥",
                    "工作时间3": "15:00", "工作内容3": "跟进明日续费事宜"}},
        {"fields": {"填写日期": "2026-07-29", "客户经理姓名": "赵六",
                    "工作时间1": "9:30", "工作内容1": "物联网卡推广，覆盖3家门店",
                    "工作时间2": "16:00", "工作内容2": "明厨亮灶对接，明日上门安装"}},
        {"fields": {"填写日期": "2026-07-29", "客户经理姓名": "孙七",
                    "工作时间1": "13:00", "工作内容1": "统付业务落地✅"}},
    ]
    stats = build_stats(today, roster, logs)
    html = generate_html(stats, "2026年07月29日（周三）")
    out = "report_selftest.html"
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[selftest] 已生成 {out}（{len(html)} 字节）")
    try:
        html_to_pdf(out, "report_selftest.pdf")
        print("[selftest] 已生成 report_selftest.pdf")
    except Exception as e:
        print(f"[selftest] PDF 跳过（未安装 playwright 或本地无 chromium）：{e}")


# ===================== 主流程 =====================
def main():
    if os.environ.get("SELFTEST") == "1":
        selftest()
        return

    app_id = os.environ.get("FEISHU_APP_ID")
    app_secret = os.environ.get("FEISHU_APP_SECRET")
    open_id = os.environ.get("FEISHU_USER_OPEN_ID")
    if not app_id or not app_secret or not open_id:
        raise SystemExit("缺少环境变量 FEISHU_APP_ID / FEISHU_APP_SECRET / FEISHU_USER_OPEN_ID")

    now = datetime.datetime.now(BJT)
    override = os.environ.get("REPORT_DATE")
    if override:
        today = override
        d = datetime.datetime.strptime(override, "%Y-%m-%d")
        today_label = d.strftime("%Y年%m月%d日") + f"（{get_weekday(d)}）"
    else:
        today = now.strftime("%Y-%m-%d")
        today_label = now.strftime("%Y年%m月%d日") + f"（{get_weekday(now)}）"

    token = get_tenant_token(app_id, app_secret)

    # 去重：若本档位今日已由另一触发器（飞书自动化 / GitHub 自带定时器）生成过，跳过防重复
    slot = get_slot(now)
    if slot is not None:
        if str(slot) in read_fired_slots(token, "report", today):
            print(f"[dedup] 档位 {slot} 今日已生成报告，跳过（防重复）")
            return

    roster_items = list_all_records(ROSTER_APP, ROSTER_TABLE, token)
    log_items = list_all_records(LOG_APP, LOG_TABLE, token)
    stats = build_stats(today, roster_items, log_items)

    html = generate_html(stats, today_label)
    html_path = f"客户经理日志报告_{today}.html"
    pdf_path = f"客户经理日志报告_{today}.pdf"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[info] HTML 已生成：{html_path}")

    html_to_pdf(html_path, pdf_path)
    print(f"[info] PDF 已生成：{pdf_path}")

    intro = (f"📈 {today_label} 客户经理日志分析报告（提交 {len(stats['submitted'])}/{stats['expected']} 人，"
             f"提交率 {stats['submit_rate']:.0f}%，总段数 {stats['total_segs']}），见附件PDF。")
    send_text_message(token, open_id, intro)
    with open(pdf_path, "rb") as f:
        file_key = upload_file(token, f.read(), pdf_path, "pdf", "application/pdf")
    resp = send_file_message(token, open_id, file_key)
    print("file 返回:", resp)
    print(f"✅ 已发送 {pdf_path}")
    if slot is not None:
        mark_slot_fired(token, "report", today, slot)


if __name__ == "__main__":
    main()
