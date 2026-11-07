# -*- coding: utf-8 -*-
"""
一人付（opcpay）接入助手 —— 供 generator_web.py 调用。

职责：
  1) create_purchase(user_id, plan_key) -> purchase_url
     调一人付 POST /purchase 创建购买入口，返回支付宝付款链接。
  2) get_entitlement(user_id) / parse_pro_entitlement(resp)
     调一人付 POST /access 查询用户当前权益（时长包生效/到期日），
     作为付费权益的权威来源；取代 webhook / Edge Function 方案。
  3) PLANS：你的三个套餐（sku 与天数），上线前把 sku 改成一人付后台实际值。

依赖：仅用标准库 urllib（Streamlit Cloud 默认可用），不依赖 requests。
API Key 从 st.secrets["OPCPAY_API_KEY"] 读取（服务端，不要暴露给前端）。
"""

import json
import logging
import os
import time
import uuid
import urllib.request
from datetime import datetime, timedelta

OPCPAY_BASE = "https://ai6666.com/opcpay/api/pay/v1"

# sku_id 取自一人付后台「收费配置」实际生成的值（已填入，勿随意改动）
PLANS = {
    "month": {"label": "月卡 ¥9.9/月", "sku": "S0806000213847679", "plan_days": 30},
    "year":  {"label": "年卡 ¥39/年",  "sku": "S0806000213847680", "plan_days": 365},
    "five":  {"label": "5年卡 ¥99/5年", "sku": "S0806000213847681", "plan_days": 1825},
}


def _api_key():
    # Streamlit 环境从 secrets 读；本地调试可走环境变量
    try:
        import streamlit as st
        return st.secrets.get("OPCPAY_API_KEY", "")
    except Exception:
        return os.environ.get("OPCPAY_API_KEY", "")


def create_purchase(user_id, plan_key, supabase=None):
    """创建购买入口，返回支付宝付款链接 purchase_url。

    user_id  : 8D 应用的登录账号（邮箱/手机号），会作为一人付的外部 userId
    plan_key : "month" / "year" / "five"（见 PLANS）
    supabase : 预留参数（历史兼容）；当前付费权益以 /access 查询为准，不再写 opcpay_orders。
    """
    if plan_key not in PLANS:
        raise ValueError("未知套餐: " + str(plan_key))
    plan = PLANS[plan_key]

    idem = f"{user_id}:{uuid.uuid4()}"

    body = {
        "user_id": user_id,
        "idempotency_key": idem,
        "sku_id": plan["sku"],
    }
    req = urllib.request.Request(
        OPCPAY_BASE + "/purchase",
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + _api_key(),
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            resp = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        _body = ""
        try:
            _body = he.read().decode("utf-8", "ignore")
        except Exception:
            pass
        logging.error("opcpay /purchase HTTP %s: %s" % (he.code, _body))
        raise RuntimeError("一人付创建订单失败(HTTP %s): %s" % (he.code, _body))
    except Exception as e:
        logging.error("opcpay /purchase 请求异常: %s" % e)
        raise

    purchase_url = resp.get("purchase_url")
    if not purchase_url:
        raise RuntimeError("一人付未返回 purchase_url，响应: " + json.dumps(resp, ensure_ascii=False))

    return purchase_url


# ---------------------------------------------------------------------------
# 权益查询（POST /access）—— 付费状态的权威来源
#
# 一人付文档明确：
#   - webhook 仅在"创建购买入口后对该用户权益持续监测 24 小时"，并非实时；
#   - "需要当前权益时重新查询 用户权益接口"；
#   - 通知不包含 payment.succeeded，不能证明某笔付款成功。
# 因此本应用采用"登录 / 生成报告时查 /access"的方式作为付费权益真相，
# 彻底免去 webhook / Supabase Edge Function / 4 个密钥等复杂环节。
# ---------------------------------------------------------------------------

_ENTITLEMENT_CACHE = {}
_ENTITLEMENT_TTL = 300  # 秒：每用户最多 5 分钟查一次一人付，避免每次页面渲染都打 API


def _fetch_entitlement(user_id):
    """实际调用 POST /access，返回原始 JSON（失败返回 None）。"""
    try:
        key = _api_key()
        if not key:
            return None
        body = {"user_id": user_id, "amount": 1}
        req = urllib.request.Request(
            OPCPAY_BASE + "/access",
            data=json.dumps(body).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": "Bearer " + key,
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:
        logging.error("opcpay /access 查询失败: %s" % e)
        return None


def get_entitlement(user_id):
    """带 TTL 缓存的权益查询；返回原始 /access 响应 dict（可能含 access_granted / decision / valid_until 等）。"""
    now = time.time()
    cached = _ENTITLEMENT_CACHE.get(user_id)
    if cached and (now - cached[1]) < _ENTITLEMENT_TTL:
        return cached[0]
    result = _fetch_entitlement(user_id)
    _ENTITLEMENT_CACHE[user_id] = (result, now)
    return result


def clear_entitlement_cache(user_id=None):
    """清除权益缓存。付款成功后调用，使下次查询立即生效。"""
    if user_id is None:
        _ENTITLEMENT_CACHE.clear()
    else:
        _ENTITLEMENT_CACHE.pop(user_id, None)


def parse_pro_entitlement(resp):
    """从 /access 响应解析"当前是否有生效的付费权益"。

    返回 (is_pro: bool, expire_iso: str|None)
      - is_pro=True,  expire_iso=None      → 永久/无到期（按 pro 处理）
      - is_pro=True,  expire_iso 有值       → 时长包，到期日（ISO 格式）
      - is_pro=False                        → 无生效权益（不覆盖任何现有 license）

    OpenAPI 字段（时长包 ALLOW 分支）：
      access_granted(bool) / decision("ALLOW") / valid_until("YYYY-MM-DD HH:MM:SS")
    """
    if not isinstance(resp, dict):
        return (False, None)
    if not resp.get("access_granted") or resp.get("decision") != "ALLOW":
        return (False, None)
    vu = resp.get("valid_until")
    if not vu:
        # ALLOW 但无 valid_until（如买断/永久），视为无期限 pro
        return (True, None)
    # valid_until 形如 "2026-11-13 10:00:00"（也可能带 T 或仅日期）
    exp = None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            exp = datetime.strptime(vu, fmt)
            break
        except Exception:
            continue
    if exp is None:
        try:
            exp = datetime.fromisoformat(vu)
        except Exception:
            return (True, None)  # 解析不出但已授权，按无期限处理
    if exp <= datetime.now():
        return (False, None)  # 已过期，不授予
    return (True, exp.isoformat())
