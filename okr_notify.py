#!/usr/bin/env python3
"""
OKR 飞书通知脚本（开放平台应用版）
每天运行，检查当天是否有 OKR 事件：
  - 上级事件 → 私发给 FEISHU_MANAGER_IDS 中的每个账号
  - 员工/团队事件 → 发送至机器人所在的所有群

环境变量：
  FEISHU_APP_ID       开放平台应用的 App ID
  FEISHU_APP_SECRET   开放平台应用的 App Secret
  FEISHU_MANAGER_IDS  上级的 open_id，多个用英文逗号分隔
  SKIP_SSL_VERIFY     本地测试遇到企业代理 SSL 报错时设为 1
"""

import os
import json
import urllib.request
from datetime import date, timedelta
from calendar import monthrange

# ── 节假日数据（国办发明电〔2025〕7号）────────────────────────────────────────

HOLIDAYS = {
    # 2025 元旦
    "2025-01-01",
    # 2025 春节
    "2025-01-28", "2025-01-29", "2025-01-30", "2025-01-31",
    "2025-02-01", "2025-02-02", "2025-02-03", "2025-02-04",
    # 2025 清明
    "2025-04-04", "2025-04-05", "2025-04-06",
    # 2025 劳动节
    "2025-05-01", "2025-05-02", "2025-05-03", "2025-05-04", "2025-05-05",
    # 2025 端午
    "2025-05-31", "2025-06-01", "2025-06-02",
    # 2025 国庆+中秋
    "2025-10-01", "2025-10-02", "2025-10-03", "2025-10-04",
    "2025-10-05", "2025-10-06", "2025-10-07", "2025-10-08",
    # 2026 元旦
    "2026-01-01", "2026-01-02", "2026-01-03",
    # 2026 春节
    "2026-02-15", "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19",
    "2026-02-20", "2026-02-21", "2026-02-22", "2026-02-23",
    # 2026 清明
    "2026-04-04", "2026-04-05", "2026-04-06",
    # 2026 劳动节
    "2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05",
    # 2026 端午
    "2026-06-19", "2026-06-20", "2026-06-21",
    # 2026 中秋
    "2026-09-25", "2026-09-26", "2026-09-27",
    # 2026 国庆
    "2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04",
    "2026-10-05", "2026-10-06", "2026-10-07",
}

# 补班日（周末当工作日）
MAKEUP = {
    # 2025
    "2025-01-26", "2025-02-08", "2025-04-27", "2025-09-28", "2025-10-11",
    # 2026
    "2026-01-04", "2026-02-14", "2026-02-28",
    "2026-05-09", "2026-09-20", "2026-10-10",
}

# ── 工作日计算 ────────────────────────────────────────────────────────────────

def is_workday(d: date) -> bool:
    s = d.isoformat()
    if s in MAKEUP:
        return True
    if s in HOLIDAYS:
        return False
    return d.weekday() < 5  # 0=周一 … 4=周五


def nth_workday(year: int, month: int, n: int) -> date | None:
    d = date(year, month, 1)
    count = 0
    while d.month == month:
        if is_workday(d):
            count += 1
            if count == n:
                return d
        d += timedelta(days=1)
    return None


def last_nth_workday(year: int, month: int, n: int) -> date | None:
    last_day = monthrange(year, month)[1]
    d = date(year, month, last_day)
    count = 0
    while d.month == month:
        if is_workday(d):
            count += 1
            if count == n:
                return d
        d -= timedelta(days=1)
    return None


def work_fridays(year: int, month: int) -> list[date]:
    result = []
    d = date(year, month, 1)
    while d.month == month:
        if d.weekday() == 4 and is_workday(d):  # 4 = 周五
            result.append(d)
        d += timedelta(days=1)
    return result

# ── 事件生成 ──────────────────────────────────────────────────────────────────

PHASE_LABEL = {"plan": "制定", "track": "跟进", "review": "复盘", "cult": "文化"}
WEEKDAY_ZH  = ["一", "二", "三", "四", "五", "六", "日"]


def get_events_for_month(year: int, month: int) -> list[dict]:
    events = []

    def add(d, label, phase, role, time=None, desc=None, link=None):
        if d:
            events.append({"date": d, "label": label, "phase": phase,
                            "role": role, "time": time, "desc": desc, "link": link})

    wd5 = nth_workday(year, month, 5)
    wd5_label = f"{wd5.month}月{wd5.day}日" if wd5 else "第5个工作日"

    add(nth_workday(year, month, 1), "上级提交文化评价问卷", "cult", "上级",
        desc="请填写对团队成员的文化评价问卷",
        link={"text": "点击填写", "url": "https://n6fo0mbcz6.feishu.cn/share/base/form/shrcnXu5NLgrWDKrcA9E34E8IFf"})

    add(nth_workday(year, month, 1), "预约上级 1v1 月度绩效沟通", "review", "员工",
        desc=f"请在本月第5个工作日（{wd5_label}）前与上级预约月度绩效沟通会议，未完成沟通将影响当月绩效工资发放")

    add(nth_workday(year, month, 2), "团队月度复盘会", "review", "团队",
        desc="请主动联系上级，预约并安排本次月度复盘会议")

    add(wd5, "上级完成上月 OKR 评分 & 确认本月 OKR", "plan", "上级", time="18:00",
        desc="与各成员进行 1v1，确认本月绩效结果")

    for fri in work_fridays(year, month):
        add(fri, "OKR 进展更新 & 周度复盘", "track", "员工", time="18:00",
            desc="在 OKR 系统上更新本周进度，并提交 OKR 进展报告")

    add(last_nth_workday(year, month, 3), "启动下月 OKR 制定 & 开始本月自评", "review", "员工",
        desc="在 OKR 系统上开始起草下月 OKR，并撰写本月 OKR 复盘文档")

    add(last_nth_workday(year, month, 1), "提交自评文档 & 下月 OKR 初稿 & 文化自评问卷", "review", "员工",
        desc="在 OKR 系统上提交本月 OKR 复盘文档，并发布下月 OKR；另需填写文化自评问卷：",
        link={"text": "点击填写", "url": "https://n6fo0mbcz6.feishu.cn/share/base/form/shrcnOUGAuZGNRck8PbwBvcHevf"})

    events.sort(key=lambda e: e["date"])
    return events


def get_todays_events() -> list[dict]:
    today = date.today()
    return [e for e in get_events_for_month(today.year, today.month)
            if e["date"] == today]

# ── 飞书开放平台 API ──────────────────────────────────────────────────────────

import ssl

BASE_URL = "https://open.feishu.cn/open-apis"


def _ssl_ctx():
    ctx = ssl.create_default_context()
    if os.environ.get("SKIP_SSL_VERIFY"):
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _request(method: str, path: str, token: str = None, payload: dict = None) -> dict:
    data = json.dumps(payload).encode() if payload else None
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(
        BASE_URL + path, data=data, headers=headers, method=method
    )
    with urllib.request.urlopen(req, timeout=10, context=_ssl_ctx()) as resp:
        return json.loads(resp.read())


def get_token(app_id: str, app_secret: str) -> str:
    body = _request("POST", "/auth/v3/tenant_access_token/internal",
                    payload={"app_id": app_id, "app_secret": app_secret})
    if body.get("code", -1) != 0:
        raise RuntimeError(f"获取 token 失败：{body}")
    return body["tenant_access_token"]


def get_bot_chat_ids(token: str) -> list[str]:
    """列出机器人所在的所有群 chat_id（自动翻页）"""
    chat_ids, page_token = [], None
    while True:
        path = "/im/v1/chats?page_size=100"
        if page_token:
            path += f"&page_token={page_token}"
        body = _request("GET", path, token=token)
        if body.get("code", -1) != 0:
            raise RuntimeError(f"获取群列表失败：{body}")
        data = body.get("data", {})
        chat_ids.extend(item["chat_id"] for item in data.get("items", []))
        if not data.get("has_more"):
            break
        page_token = data.get("page_token")
    return chat_ids


def send_message(token: str, receive_id_type: str, receive_id: str,
                 events: list[dict], today: date) -> None:
    weekday = WEEKDAY_ZH[today.weekday()]
    lines = []
    lines.append([{"tag": "text",
                   "text": f"今天共有 {len(events)} 项 OKR 事务，请相关同学及时处理："}])
    lines.append([{"tag": "text", "text": ""}])
    for ev in events:
        time_text = f"  |  {ev['time']} 前" if ev["time"] else ""
        lines.append([{"tag": "text", "text": f"【{PHASE_LABEL[ev['phase']]}】{ev['label']}{time_text}"}])
        if ev.get("desc"):
            lines.append([{"tag": "text", "text": f"  {ev['desc']}"}])
        if ev.get("link"):
            lines.append([{"tag": "a", "text": f"  {ev['link']['url']}", "href": ev["link"]["url"]}])
        lines.append([{"tag": "text", "text": ""}])

    content_obj = {
        "zh_cn": {
            "title": f"📅 OKR 今日提醒 · {today.month}月{today.day}日 周{weekday}",
            "content": lines,
        }
    }
    payload = {
        "receive_id": receive_id,
        "msg_type": "post",
        "content": json.dumps(content_obj),
    }
    body = _request("POST", f"/im/v1/messages?receive_id_type={receive_id_type}",
                    token=token, payload=payload)
    if body.get("code", -1) != 0:
        raise RuntimeError(f"发送消息失败（{receive_id}）：{body}")


# ── 入口 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app_id      = os.environ.get("FEISHU_APP_ID", "").strip()
    app_secret  = os.environ.get("FEISHU_APP_SECRET", "").strip()
    manager_ids = [x.strip() for x in
                   os.environ.get("FEISHU_MANAGER_IDS", "").split(",") if x.strip()]

    if not app_id or not app_secret:
        raise SystemExit("❌ 未设置 FEISHU_APP_ID 或 FEISHU_APP_SECRET")

    today = date.today()
    all_events = get_todays_events()

    if not all_events:
        print(f"✅ {today}：今天没有 OKR 事务，不发通知")
        raise SystemExit(0)

    token = get_token(app_id, app_secret)

    # 上级事件 → 私信
    manager_events = [e for e in all_events if e["role"] == "上级"]
    if manager_events:
        if not manager_ids:
            print("⚠️  有上级事务但未配置 FEISHU_MANAGER_IDS，跳过私信")
        else:
            for uid in manager_ids:
                send_message(token, "open_id", uid, manager_events, today)
            print(f"✅ 上级私信 → {len(manager_ids)} 人，共 {len(manager_events)} 项事务")

    # 员工/团队事件 → 群发
    group_events = [e for e in all_events if e["role"] != "上级"]
    if group_events:
        chat_ids = get_bot_chat_ids(token)
        if not chat_ids:
            print("⚠️  机器人未加入任何群，跳过群发")
        else:
            for chat_id in chat_ids:
                send_message(token, "chat_id", chat_id, group_events, today)
            print(f"✅ 群消息 → {len(chat_ids)} 个群，共 {len(group_events)} 项事务")
