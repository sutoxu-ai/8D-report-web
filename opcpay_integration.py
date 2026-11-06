# -*- coding: utf-8 -*-
"""
一人付（opcpay）接入助手 —— 供 generator_web.py 调用。

职责：
  1) create_purchase(user_id, sku, plan_days, supabase) -> purchase_url
     调一人付 POST /purchase 创建购买入口，返回支付宝付款链接；
     同时在 Supabase 的 opcpay_orders 表写入"待激活订单"，供 webhook 映射有效期。
  2) PLANS：你的三个套餐（sku 与天数），上线前把 sku 改成一人付后台实际值。

依赖：仅用标准库 urllib（Streamlit Cloud 默认可用），不依赖 requests。
API Key 从 st.secrets["OPCPAY_API_KEY"] 读取（服务端，不要暴露给前端）。
"""

import json
import os
import uuid
import urllib.request
from datetime import datetime, timedelta

OPCPAY_BASE = "https://ai6666.com/opcpay/api/pay/v1"

# ⚠️ 上线前把下面三个 sku 改成你在一人付后台「收费配置」里实际生成的 skuId
PLANS = {
    "month": {"label": "月卡 ¥9.9/月", "sku": "REPLACE_MONTH_SKU", "plan_days": 30},
    "year":  {"label": "年卡 ¥39/年",  "sku": "REPLACE_YEAR_SKU",  "plan_days": 365},
    "five":  {"label": "5年卡 ¥99/5年", "sku": "REPLACE_FIVE_SKU",  "plan_days": 1825},
}


def _api_key():
    # Streamlit 环境从 secrets 读；本地调试可走环境变量
    try:
        import streamlit as st
        return st.secrets.get("OPCPAY_API_KEY", "")
    except Exception:
        return os.environ.get("OPCPAY_API_KEY", "")


def create_purchase(user_id, plan_key, supabase):
    """创建购买入口，返回支付宝付款链接 purchase_url。

    user_id  : 8D 应用的登录账号（邮箱/手机号），会作为一人付的外部 userId
    plan_key : "month" / "year" / "five"（见 PLANS）
    supabase : 已初始化的 supabase 客户端（用于写 opcpay_orders）
    """
    if plan_key not in PLANS:
        raise ValueError("未知套餐: " + str(plan_key))
    plan = PLANS[plan_key]

    idem = f"{user_id}:{uuid.uuid4()}"

    # 先写待激活订单：webhook 会据此映射有效期。
    # 放在调用 opcpay 之前，确保 DB 异常尽早暴露（不会让用户拿到付款链接却无法激活）。
    supabase.table("opcpay_orders").insert({
        "user_id": user_id,
        "idempotency_key": idem,
        "sku": plan["sku"],
        "plan_days": plan["plan_days"],
        "fulfilled": False,
    }).execute()

    body = {
        "userId": user_id,
        "idempotencyKey": idem,
        "skuId": plan["sku"],
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
    with urllib.request.urlopen(req, timeout=30) as r:
        resp = json.loads(r.read().decode("utf-8"))

    purchase_url = resp.get("purchase_url")
    if not purchase_url:
        raise RuntimeError("一人付未返回 purchase_url，响应: " + json.dumps(resp, ensure_ascii=False))

    return purchase_url
