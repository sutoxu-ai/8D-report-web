#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
8D 报告智能生成助手 - 客户端
最终优化版本
"""

import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO
from datetime import datetime, timedelta
import re
import base64
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import openai
from supabase import create_client
import logging

# ==================== 缓存配置 ====================
@st.cache_data(ttl=60)
def get_cached_license(user_id):
    """缓存用户许可证信息，60 秒 TTL"""
    if not supabase:
        return None
    try:
        r = supabase.table("licenses").select("*").eq("user_id", user_id).execute()
        if r.data:
            return r.data[0]
        return create_free_license(user_id)
    except Exception:
        return None

def clear_license_cache(user_id):
    """清除特定用户的缓存"""
    get_cached_license.clear()

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="8D 报告 - 智能生成助手",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 隐藏 Streamlit 默认 UI 元素 ====================
hide_streamlit_style = """
<style>
    /* 只隐藏右上角菜单 */
    #MainMenu {visibility: hidden !important; display: none !important;}

    /* ========== 顶部 header / 工具栏改为蓝色背景 ========== */
    header[data-testid="stHeader"],
    [data-testid="stToolbar"] {
        background-color: #1e3a5f !important;
        color: #ffffff !important;
    }
    header[data-testid="stHeader"] *,
    [data-testid="stToolbar"] * {
        color: #ffffff !important;
    }
    header[data-testid="stHeader"] a,
    [data-testid="stToolbar"] a {
        color: #ffffff !important;
    }
    /* Deploy 按钮：真实 testid=stAppDeployButton / class=stAppDeployButton
       文字 + 背景 + 边框 都与 header 同色 #1e3a5f，肉眼看不见 */
    .stAppDeployButton,
    [data-testid="stAppDeployButton"],
    [data-testid="stToolbar"] .stAppDeployButton,
    header[data-testid="stHeader"] .stAppDeployButton {
        color: #1e3a5f !important;
        background: #1e3a5f !important;
        background-color: #1e3a5f !important;
        background-image: none !important;
        border: 1px solid #1e3a5f !important;
        box-shadow: none !important;
    }
    /* 按钮内部文字 / 图标也一起同色 */
    .stAppDeployButton *,
    [data-testid="stAppDeployButton"] * {
        color: #1e3a5f !important;
        fill: #1e3a5f !important;
        stroke: #1e3a5f !important;
        background: #1e3a5f !important;
        background-color: #1e3a5f !important;
    }

    /* 工具栏里的所有 action 按钮（Stop / Share / ★ / ✏ / ⟳ 等）整体与 header 同色 */
    [data-testid="stAppToolbar"] .stToolbarActionButton,
    .stAppToolbar .stToolbarActionButton,
    .stToolbarActions button,
    [data-testid="stAppToolbar"] [data-testid="stAppShareButton"],
    [data-testid="stAppToolbar"] [data-testid="stAppStopButton"],
    [data-testid="stAppShareButton"],
    [data-testid="stAppStopButton"] {
        color: #1e3a5f !important;
        background: #1e3a5f !important;
        background-color: #1e3a5f !important;
        background-image: none !important;
        border: 1px solid #1e3a5f !important;
        border-color: #1e3a5f !important;
        box-shadow: none !important;
    }
    .stAppToolbar .stToolbarActionButton *,
    [data-testid="stAppToolbar"] .stToolbarActionButton *,
    [data-testid="stAppToolbar"] [data-testid="stAppShareButton"] *,
    [data-testid="stAppToolbar"] [data-testid="stAppStopButton"] * {
        color: #1e3a5f !important;
        fill: #1e3a5f !important;
        stroke: #1e3a5f !important;
        background: #1e3a5f !important;
        background-color: #1e3a5f !important;
    }

    /* 隐藏 footer 水印 */
    footer {visibility: hidden !important; display: none !important;}
    
    /* 隐藏 Pages 导航菜单列表 */
    [data-testid="stSidebarNav"] > ul {display: none !important;}
    
    /* 调整主内容区域 */
    .main .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
    }
    /* 顶部留给 header 的空间尽量小 */
    .stApp > header[data-testid="stHeader"] {
        height: 2.2rem !important;
        min-height: 2.2rem !important;
    }
    
    /* ========== 缩小侧边栏间距 ========== */
    [data-testid="stSidebar"] .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
    [data-testid="stSidebar"] h3 {
        margin-top: 0.5rem !important;
        margin-bottom: 0.3rem !important;
    }
    
    [data-testid="stSidebar"] p {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    [data-testid="stSidebar"] .stButton {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    [data-testid="stSidebar"] .stTextInput {
        margin-top: 0.2rem !important;
        margin-bottom: 0.2rem !important;
    }
    
    [data-testid="stSidebar"] .streamlit-expanderHeader {
        padding-top: 0.3rem !important;
        padding-bottom: 0.3rem !important;
    }
    
    [data-testid="stSidebar"] .stCaption {
        margin-top: 0.1rem !important;
        margin-bottom: 0.1rem !important;
    }
    
    [data-testid="stSidebar"] .stMarkdown {
        margin-bottom: 0.3rem !important;
    }
    
    [data-testid="stSidebar"] div[data-testid="stVerticalBlock"] > div {
        gap: 0.15rem !important;
    }
    
    [data-testid="stSidebar"] hr {
        display: none !important;
    }
    
    /* ========== 手机端字体缩小 ========== */
    @media screen and (max-width: 768px) {
        h1 { font-size: 1.3rem !important; }
        h2 { font-size: 1.1rem !important; }
        h3 { font-size: 1rem !important; }
        h4 { font-size: 0.95rem !important; }
        body { font-size: 0.85rem !important; }
        input, textarea { font-size: 0.85rem !important; }
        button { font-size: 0.9rem !important; }
        label, .stMarkdown p, .stMarkdown span { font-size: 0.85rem !important; }
        /* 手机端缩小文本输入框高度 */
        textarea {
            height: 80px !important;
            min-height: 80px !important;
        }
    }
    
    /* 品牌头部样式见下方（深色画布上的实色条） */
    @media screen and (max-width: 768px) {
        .brand-logo { width: 38px; height: 38px; font-size: 1rem; }
        .brand-title { font-size: 1.1rem; }
        .brand-subtitle { font-size: 0.7rem; }
    }
    
    /* ========== 输入卡片样式（深蓝底，配合白色文字） ========== */
    div[data-testid="stVerticalBlock"] .input-card {
        border: 1px solid #2a3654;
        border-radius: 0.6rem;
        overflow: hidden;
        margin-bottom: 0.4rem;
        background: #131a2c;
        box-shadow: 0 2px 8px rgba(0,0,0,0.25);
    }
    .input-card-body {
        padding: 0.4rem 0.6rem;
    }

    /* ========== 主区紧凑布局：缩小卡片间距、表单元素间距 ========== */
    /* ★ 不依赖 .main 类名（不同 Streamlit 版本结构有变），
       直接锁 data-testid，只要在主区里出现就生效 */
    [data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"],
    .main [data-testid="stVerticalBlockBorderWrapper"] {
        margin: 0 !important;
        padding: 0.1rem 0.7rem !important;
        border-top-width: 1px !important;
    }
    /* ★ 关键：主区垂直栈的内部 gap，彻底压为 0 */
    [data-testid="stMain"] [data-testid="stVerticalBlock"] > div,
    .main [data-testid="stVerticalBlock"] > div {
        gap: 0 !important;
    }
    /* 元素容器外边距压到最小 */
    [data-testid="stMain"] .element-container,
    .main .element-container {
        margin: 0 !important;
    }
    /* 输入控件整体压矮 */
    .main [data-testid="stTextInput"] input,
    .main [data-testid="stNumberInput"] input,
    .main [data-testid="stDateInput"] input,
    .main [data-testid="stSelectbox"] [data-testid="stWidgetCombobox"] {
        min-height: 32px !important;
        padding: 0.15rem 0.5rem !important;
        font-size: 0.85rem !important;
    }
    .main [data-testid="stTextArea"] textarea {
        min-height: 60px !important;
        padding: 0.2rem 0.5rem !important;
        font-size: 0.85rem !important;
    }
    /* selectbox/date 内部触发按钮也压矮 */
    .main [data-baseweb="select"] > div,
    .main [data-baseweb="input"] > div {
        min-height: 32px !important;
    }
    /* label 与输入框之间：去 margin，强制贴近 */
    .main label, .main [data-testid="stWidgetLabel"] {
        margin: 0 0 0.05rem 0 !important;
        padding: 0 !important;
        font-size: 0.82rem !important;
        line-height: 1.2 !important;
    }
    /* 输入框根容器：去掉顶部留白 */
    .main [data-testid="stTextInput"],
    .main [data-testid="stTextArea"],
    .main [data-testid="stNumberInput"],
    .main [data-testid="stDateInput"],
    .main [data-testid="stSelectbox"],
    .main [data-testid="stMultiSelect"] {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }
    /* ★ 关键：Streamlit 把 label 和 input 包在一个 flex 列容器里，
       默认 gap 很大（约 0.5rem）。直接压这个 gap 才能把两者拉近。 */
    .main [data-testid="stTextInput"] > div,
    .main [data-testid="stTextArea"] > div,
    .main [data-testid="stNumberInput"] > div,
    .main [data-testid="stDateInput"] > div,
    .main [data-testid="stSelectbox"] > div,
    .main [data-testid="stMultiSelect"] > div,
    .main [data-testid="stTextInput"] > div > div,
    .main [data-testid="stTextArea"] > div > div,
    .main [data-testid="stNumberInput"] > div > div,
    .main [data-testid="stDateInput"] > div > div,
    .main [data-testid="stSelectbox"] > div > div,
    .main [data-testid="stMultiSelect"] > div > div {
        display: flex !important;
        flex-direction: column !important;
        gap: 0.05rem !important;
        margin-top: 0 !important;
        padding-top: 0 !important;
    }
    /* Streamlit 给 label 和 input 中间塞的 st-emotion 容器，去 padding */
    .main [data-testid="stWidgetLabel"] + div,
    .main label + [data-baseweb="input"],
    .main label + [data-baseweb="select"],
    .main label + [data-baseweb="textarea"] {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }

    /* ========== 报告预览区所有文字（白色，深蓝底上看得清） ========== */
    /* 直接用 class/data-testid 锁定，不依赖 .main 类名 */
    [data-testid="stMarkdown"],
    [data-testid="stMarkdownContainer"],
    .report-preamble,
    .d-section-body,
    .d-section-title,
    .d-section {
        color: #ffffff !important;
    }
    [data-testid="stMarkdown"] *,
    [data-testid="stMarkdownContainer"] *,
    .report-preamble *,
    .d-section-body *,
    .d-section-body p,
    .d-section-body li,
    .d-section-body strong,
    .d-section-body span,
    .d-section-body table,
    .d-section-body table th,
    .d-section-body table td,
    .d-section-body h1, .d-section-body h2, .d-section-body h3,
    .d-section-body h4, .d-section-body h5, .d-section-body h6 {
        color: #ffffff !important;
    }
    .d-section-body table th {
        background: #1f2a44 !important;
        color: #ffffff !important;
    }
    .report-preamble {
        background: transparent !important;
    }

    /* ========== 输入卡片内表单标签（白色，深底上才看得清） ========== */
    /* 用 html body 前缀拉高特异性，再叠加 !important，必杀 */
    html body [data-testid="stWidgetLabel"],
    html body label[data-testid="stWidgetLabel"],
    html body [data-testid="stTextInput"] label,
    html body [data-testid="stTextInput"] [data-testid="stWidgetLabel"],
    html body [data-testid="stTextArea"] label,
    html body [data-testid="stTextArea"] [data-testid="stWidgetLabel"],
    html body [data-testid="stNumberInput"] label,
    html body [data-testid="stNumberInput"] [data-testid="stWidgetLabel"],
    html body [data-testid="stDateInput"] label,
    html body [data-testid="stDateInput"] [data-testid="stWidgetLabel"],
    html body [data-testid="stSelectbox"] label,
    html body [data-testid="stSelectbox"] [data-testid="stWidgetLabel"],
    html body [data-testid="stMultiselect"] label,
    html body [data-testid="stMultiselect"] [data-testid="stWidgetLabel"],
    .main label,
    .main .stTextInput label,
    .main .stTextInput > label,
    .main .stTextArea label,
    .main .stTextArea > label,
    .main .stNumberInput label,
    .main .stDateInput label,
    .main .stSelectbox label,
    .main .stMultiselect label,
    .main [data-testid="stWidgetLabel"] {
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    /* placeholder 也要看得清 */
    .main input::placeholder,
    .main textarea::placeholder {
        color: #64748b !important;
    }

    /* ========== 侧边栏文字（白色，深蓝底上清楚） ========== */
    /* 输入激活码、生成历史、按钮等 */
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
    [data-testid="stSidebar"] .streamlit-expanderHeader {
        color: #ffffff !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdown"] h1,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] h2,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] h3,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] h4,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] li,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] span,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] strong,
    [data-testid="stSidebar"] [data-testid="stMarkdown"] em {
        color: #ffffff !important;
    }
    /* 侧边栏内的成功/错误/提示信息也白 */
    [data-testid="stSidebar"] .stAlert,
    [data-testid="stSidebar"] .stAlert * {
        color: #ffffff !important;
    }

    /* ========== st.status 生成中 / 完成 文字（白色） ========== */
    /* 含“✅ 报告生成完成！”标签 与 正在流式输出的正文 */
    [data-testid="stStatus"],
    [data-testid="stStatus"] *,
    [data-testid="stStatus"] .stStatusLabel,
    [data-testid="stStatus"] [data-testid="stMarkdown"],
    [data-testid="stStatus"] [data-testid="stMarkdownContainer"],
    [data-testid="stStatus"] [data-testid="stMarkdown"] *,
    [data-testid="stStatus"] [data-testid="stMarkdownContainer"] *,
    [data-testid="stStatus"] label,
    [data-testid="stStatus"] p,
    [data-testid="stStatus"] li,
    [data-testid="stStatus"] span,
    [data-testid="stStatus"] strong,
    [data-testid="stStatus"] div,
    [data-testid="stStatus"] table,
    [data-testid="stStatus"] td,
    [data-testid="stStatus"] th,
    [data-testid="stStatus"] .streamlit-expanderHeader {
        color: #ffffff !important;
    }
    [data-testid="stStatus"] {
        background: transparent !important;
        border-color: #2a3a5c !important;
    }
    
    /* ========== D 章节彩色边框（深底） ========== */
    .d-section {
        border-left: 4px solid #cbd5e1;
        padding: 0.5rem 0.8rem;
        margin-bottom: 0.6rem;
        border-radius: 0 0.4rem 0.4rem 0;
        background: #131a2c;
        color: #e2e8f0;
    }
    .d-section-d1 { border-left-color: #3b82f6; }
    .d-section-d2 { border-left-color: #22c55e; }
    .d-section-d3 { border-left-color: #f97316; }
    .d-section-d4 { border-left-color: #ec4899; }
    .d-section-d5 { border-left-color: #a855f7; }
    .d-section-d6 { border-left-color: #06b6d4; }
    .d-section-d7 { border-left-color: #eab308; }
    .d-section-d8 { border-left-color: #94a3b8; }
    .d-section-map { border-left-color: #475569; background: #1a2238; }
    .d-section-map .d-section-title { color: #94a3b8; }
    .d-section-title {
        font-weight: 700;
        font-size: 0.95rem;
        margin-bottom: 0.3rem;
    }
    .d-section-title { color: #ffffff !important; }
    .d-section-d1 .d-section-title { color: #ffffff !important; }
    .d-section-d2 .d-section-title { color: #ffffff !important; }
    .d-section-d3 .d-section-title { color: #ffffff !important; }
    .d-section-d4 .d-section-title { color: #ffffff !important; }
    .d-section-d5 .d-section-title { color: #ffffff !important; }
    .d-section-d6 .d-section-title { color: #ffffff !important; }
    .d-section-d7 .d-section-title { color: #ffffff !important; }
    .d-section-d8 .d-section-title { color: #ffffff !important; }
    .d-section-map .d-section-title { color: #ffffff !important; }
    .d-section-body {
        font-size: 0.85rem;
        color: #ffffff !important;
        line-height: 1.5;
    }
    .d-section-body *,
    .d-section-body p,
    .d-section-body li,
    .d-section-body strong,
    .d-section-body span,
    .d-section-body table,
    .d-section-body table th,
    .d-section-body table td {
        color: #ffffff !important;
    }
    /* 正文中若残留 h1-h6，强制压成正文字号。
       多重保险：双类选择器 (0,2,0) + 属性选择器 + !important，稳压 Streamlit 后加载的 .stMarkdown h1。 */
    .d-section-body.d-section-body h1, .d-section-body.d-section-body h2,
    .d-section-body.d-section-body h3, .d-section-body.d-section-body h4,
    .d-section-body.d-section-body h5, .d-section-body.d-section-body h6 {
        font-size: 0.85rem !important;
        font-weight: normal !important;
        line-height: 1.5 !important;
        margin: 0 !important;
        color: inherit !important;
        padding: 0 !important;
        border: none !important;
    }
    /* 保险栓 2：scope 化 —— 直接命中任何包含 markdown 的容器，避免 Streamlit 包装层干扰 */
    section[data-testid="stMarkdownContainer"] .d-section-body h1,
    section[data-testid="stMarkdownContainer"] .d-section-body h2,
    section[data-testid="stMarkdownContainer"] .d-section-body h3,
    section[data-testid="stMarkdownContainer"] .d-section-body h4,
    section[data-testid="stMarkdownContainer"] .d-section-body h5,
    section[data-testid="stMarkdownContainer"] .d-section-body h6 {
        font-size: 0.85rem !important;
        font-weight: normal !important;
        line-height: 1.5 !important;
        margin: 0 !important;
        color: inherit !important;
    }
    /* 保险栓 3：拦截 horizontal rule <hr> —— 章节体内不应该出现分页线 */
    .d-section-body hr {
        display: none !important;
    }
    .d-section-body table {
        border-collapse: collapse;
        width: 100%;
        margin: 0.4rem 0;
        font-size: 0.82rem;
        line-height: 1.4;
    }
    .d-section-body table th,
    .d-section-body table td {
        border: 1px solid #2a3654;
        padding: 0.3rem 0.5rem;
        text-align: left;
        white-space: pre-wrap;
        color: #e2e8f0;
    }
    .d-section-body table th {
        background: #1e293b;
        font-weight: 700;
        color: #f1f5f9;
    }
    
    /* ========== 进度圆点 ========== */
    .progress-dots {
        display: flex;
        gap: 0.3rem;
        align-items: center;
        justify-content: center;
        padding: 0.5rem 0;
    }
    .progress-dot {
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.65rem;
        font-weight: 700;
        color: white;
        transition: all 0.3s;
    }
    .dot-done { background: #22c55e; }
    .dot-active { background: #2563eb; animation: pulse 1.2s infinite; }
    .dot-pending { background: #cbd5e1; color: #94a3b8; }
    @keyframes pulse {
        0%, 100% { transform: scale(1); box-shadow: 0 0 0 0 rgba(37,99,235,0.4); }
        50% { transform: scale(1.15); box-shadow: 0 0 0 6px rgba(37,99,235,0); }
    }
    
    /* ========== 双色画布：左侧语言/账户区深蓝，右侧主区稍浅蓝 ========== */
    .stApp {
        background: #16233b;
    }
    /* 左侧：语言 / 账户区 深蓝 */
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] > div,
    [data-testid="stSidebarUserContent"],
    section[data-testid="stSidebar"] {
        background: #0b1220 !important;
    }
    /* 左侧文字改浅色，深蓝底上才看得清 */
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] .stMarkdown span,
    [data-testid="stSidebar"] .stCaption,
    [data-testid="stSidebar"] label {
        color: #cbd5e1 !important;
    }
    [data-testid="stSidebar"] h3 { color: #f1f5f9 !important; }
/* 侧边栏内输入框保持浅底深字，保证可输入可读 */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        background-color: #ffffff !important;
        color: #1e293b !important;
    }

    /* ========== 侧边栏按钮（登录、激活、加载历史、删除、退出等）统一深蓝白字 ========== */
    [data-testid="stSidebar"] button,
    [data-testid="stSidebar"] [data-testid="stButton"] button {
        background-color: #1e3a5f !important;
        background-image: none !important;
        color: #ffffff !important;
        border: 1px solid #3b82f6 !important;
    }
    [data-testid="stSidebar"] button:hover,
    [data-testid="stSidebar"] [data-testid="stButton"] button:hover {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-color: #60a5fa !important;
    }

    /* ========== 侧边栏折叠/展开按钮：按状态设图标颜色 ========== */
    /* 侧边栏打开时（按钮在内，显示 “«”）：图标始终白色 */
    [data-testid="stSidebarCollapseButton"] {
        color: #ffffff !important;
        background-color: #1e3a5f !important;
        border: 1px solid #3b82f6 !important;
        border-radius: 0.4rem !important;
        box-shadow: 0 0 0 2px rgba(59,130,246,0.35) !important;
        opacity: 1 !important;
    }
    [data-testid="stSidebarCollapseButton"] *,
    [data-testid="stSidebarCollapseButton"] svg,
    [data-testid="stSidebarCollapseButton"] svg *,
    [data-testid="stSidebarCollapseButton"] svg path,
    [data-testid="stSidebarCollapseButton"] svg g,
    [data-testid="stSidebarCollapseButton"] svg polygon,
    [data-testid="stSidebarCollapseButton"] svg line,
    [data-testid="stSidebarCollapseButton"] svg circle,
    [data-testid="stSidebarCollapseButton"] svg rect {
        fill: #ffffff !important;
        stroke: #ffffff !important;
        color: #ffffff !important;
        opacity: 1 !important;
    }
    /* 侧边栏收起时（按钮浮动在主区，显示 “»”）：图标始终黑色 */
    [data-testid="collapsedControlButton"] {
        color: #000000 !important;
        background-color: #ffffff !important;
        border: 1px solid #94a3b8 !important;
        border-radius: 0.4rem !important;
    }
    [data-testid="collapsedControlButton"] *,
    [data-testid="collapsedControlButton"] svg,
    [data-testid="collapsedControlButton"] svg *,
    [data-testid="collapsedControlButton"] svg path,
    [data-testid="collapsedControlButton"] svg g,
    [data-testid="collapsedControlButton"] svg polygon,
    [data-testid="collapsedControlButton"] svg line,
    [data-testid="collapsedControlButton"] svg circle,
    [data-testid="collapsedControlButton"] svg rect {
        fill: #000000 !important;
        stroke: #000000 !important;
        color: #000000 !important;
    }

    /* 右侧：主内容区稍浅蓝（容器透明，透出 .stApp 背景） */
    .main .block-container {
        background: transparent;
    }
    html, body, [class*="css"] {
        font-family: "PingFang SC", "Microsoft YaHei", "Hiragino Sans GB", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
    }

    /* 侧边栏保持 Streamlit 默认浅色主题 */

    /* ========== Streamlit 表单控件 / 提示框（保持默认浅色主题） ========== */

    /* ========== 品牌头部（深色画布上的实色条，保持可见） ========== */
    .brand-header {
        display: flex;
        align-items: center;
        gap: 0.8rem;
        padding: 0.5rem 1rem;
        border-radius: 0.6rem;
        background: #1e3a5f;
        margin: 0 0 0.6rem 0;
    }
    .brand-logo {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        background: rgba(255,255,255,0.15);
        color: #fff;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.2rem;
        font-weight: 800;
        flex-shrink: 0;
    }
    .brand-title {
        font-size: 1.35rem;
        font-weight: 800;
        color: #ffffff;
        margin: 0;
        line-height: 1.2;
    }
    .brand-subtitle {
        font-size: 0.78rem;
        color: rgba(255,255,255,0.85);
        margin: 0;
        line-height: 1.2;
    }

    /* ========== 面板标题 ========== */
    .panel-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: #f1f5f9;
        margin: 0 0 0.4rem 0;
        padding-bottom: 0.3rem;
        border-bottom: 1px solid #2a3654;
        letter-spacing: 0.3px;
    }

    /* ========== 输入卡片头部（彩色标签）—— 浅色主题 ========== */
    .input-card-header {
        display: inline-block;
        margin: 0 0 0.7rem 0;
        padding: 0.42rem 0.7rem;
        border-radius: 0.5rem;
        font-weight: 700;
        font-size: 0.86rem;
        border: none;
        border-left: 4px solid #2563eb;
        background: #eff6ff;
        color: #1e40af;
    }
    .input-card-header.accent-blue   { border-left-color: #2563eb; background: #eff6ff; color: #1e40af; }
    .input-card-header.accent-orange { border-left-color: #ea580c; background: #fff7ed; color: #c2410c; }
    .input-card-header.accent-teal   { border-left-color: #0891b2; background: #ecfeff; color: #0e7490; }

    /* ========== 输入区段间细分隔线（替代删除后的卡头） ========== */
    .form-divider {
        height: 1px;
        background: #2a3654;
        margin: 0.4rem 0;
        border: none;
        padding: 0;
    }

    /* ========== 输入卡片本体（深色主题） ========== */
    div[data-testid="stVerticalBlock"] .input-card {
        border: 1px solid #2a3654;
        border-radius: 0.6rem;
        overflow: hidden;
        margin-bottom: 0.8rem;
        background: #131a2c;
        box-shadow: 0 2px 8px rgba(0,0,0,0.25);
    }
    /* st.container(border=True) 的实际边框容器 */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #131a2c !important;
        border: 1px solid #2a3654 !important;
        border-radius: 0.6rem !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.25) !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"] > div { background-color: transparent !important; }

    /* d-section 深色样式见上方定义 */

    /* ========== D0 前置诊断卡片（深色主题） ========== */
    .d0-card {
        border: 1px solid #2a3654;
        border-radius: 0.7rem;
        padding: 0.9rem 1rem;
        margin: 0.5rem 0 0.3rem 0;
        background: linear-gradient(180deg, #131a2c, #0f1626);
        box-shadow: 0 2px 10px rgba(0,0,0,0.25);
    }
    .d0-card-header {
        font-weight: 800;
        font-size: 0.95rem;
        color: #f1f5f9;
        margin-bottom: 0.6rem;
        padding-bottom: 0.4rem;
        border-bottom: 2px dashed #2a3654;
    }
    .d0-row {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        font-size: 0.85rem;
        margin: 0.4rem 0;
    }
    .d0-row > span {
        color: #94a3b8;
        min-width: 6rem;
        flex-shrink: 0;
    }
    .d0-row > b { color: #f1f5f9; }
    .d0-bar {
        flex: 1;
        height: 9px;
        background: #1f2a44;
        border-radius: 99px;
        overflow: hidden;
    }
    .d0-bar-fill {
        height: 100%;
        border-radius: 99px;
        transition: width .4s ease;
    }
    .d0-tip {
        margin-top: 0.6rem;
        padding: 0.5rem 0.7rem;
        background: #2a1810;
        border-left: 3px solid #f59e0b;
        border-radius: 0 0.4rem 0.4rem 0;
        font-size: 0.8rem;
        color: #fbbf24;
        line-height: 1.4;
    }

    /* ========== 输入框本体（深色卡片里的白色输入框） ========== */
    .main input[type="text"],
    .main input[type="password"],
    .main textarea,
    .main [data-baseweb="input"] input,
    .main [data-baseweb="textarea"] textarea {
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #cbd5e1 !important;
    }

    /* ========== 生成按钮 ========== */
    button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb, #4338ca) !important;
        border: none !important;
        border-radius: 0.6rem !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 14px rgba(37,99,235,0.35) !important;
        transition: all .2s !important;
    }
    button[kind="primary"]:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(37,99,235,0.45) !important;
    }

    /* ========== 下载按钮（与「一键复制报告」一致的紫蓝渐变） ========== */
    .stDownloadButton button,
    [data-testid="stMain"] .stDownloadButton button,
    [data-testid="stMain"] [data-testid="baseButton-secondary"] {
        border-radius: 0.6rem !important;
        font-weight: 600 !important;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        background-color: #667eea !important;
        background-image: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: #ffffff !important;
        border: none !important;
        transition: all .2s !important;
        box-shadow: 0 2px 6px rgba(102,126,234,0.35) !important;
    }
    .stDownloadButton button:hover,
    [data-testid="stMain"] .stDownloadButton button:hover {
        background: linear-gradient(135deg, #5566d8 0%, #653a92 100%) !important;
        background-color: #5566d8 !important;
        background-image: linear-gradient(135deg, #5566d8 0%, #653a92 100%) !important;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(102,126,234,0.5) !important;
        transform: translateY(-1px);
    }
    /* 下载按钮内的 SVG 图标保持白色 */
    .stDownloadButton button svg,
    [data-testid="stMain"] .stDownloadButton button svg {
        fill: #ffffff !important;
        color: #ffffff !important;
    }

    /* ========== 普通按钮（次级，浅色主题） ========== */
    .stButton button:not([kind="primary"]) {
        border: 1px solid #cbd5e1 !important;
    }
    .stButton button:not([kind="primary"]):hover {
        border-color: #2563eb !important;
        color: #2563eb !important;
    }

    /* ========== 进度圆点（浅色微调） ========== */
    .progress-dot { color: #1e293b; }
    .dot-pending { background: #cbd5e1; color: #94a3b8; }

    /* ========== 按钮行 ========== */
    .btn-row {
        display: flex;
        gap: 0.5rem;
        margin-bottom: 0.5rem;
    }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==================== JavaScript 隐藏右上角按钮 ====================
hide_buttons_script = """
<script>
// Hide top-right toolbar buttons, keep sidebar toggle
function hideTopRightButtons() {
    // Hide links containing fork, github, star
    document.querySelectorAll('a').forEach(el => {
        const href = (el.href || '').toLowerCase();
        const text = (el.textContent || '').toLowerCase();
        if (href.includes('fork') || href.includes('github') || 
            text.includes('fork') || text.includes('star')) {
            el.style.display = 'none';
        }
    });
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(hideTopRightButtons, 500));
} else {
    setTimeout(hideTopRightButtons, 500);
}
</script>
"""
st.markdown(hide_buttons_script, unsafe_allow_html=True)

# ==================== 多语言文本 ====================
TEXT = {
    "zh": {
        "lang_label": "语言", "lang_zh": "中文", "lang_en": "English",
        "system_status": "系统状态", "pro_version": "✅ 正式版", "free_version": "⚠️ 未激活",
        "license_valid_until": "📅 有效期至 {exp}", "trial_used": "📊 已使用 {used} 次 / 共 {total} 次",
        "no_license": "⚠️ 未激活，请购买激活码", "activate_title": "🔑 授权 / 续费",
        "activate_code_hint": "激活码", "activate_btn": "立即激活",
        "activate_success": "✅ 激活成功，有效期一年", "activate_fail": "❌ 激活码无效",
        "invalid_activate_code": "请输入有效的激活码",
        "license_expired": "❌ 授权已过期",
        "login_required": "🔒 请先登录", "logout": "退出登录",
        "login_header": "👤 用户登录",
        "username_label": "邮箱或手机号",
        "username_placeholder": "例：zhangsan@163.com 或 13812345678",
        "login_register_btn": "🔓 登录 / 注册",
        "enter_username_error": "请输入邮箱或手机号",
        "invalid_email": "❌ 邮箱格式不正确，示例：zhangsan@163.com",
        "invalid_phone": "❌ 手机号格式不正确，请输入11位大陆手机号（1开头，第二位3-9）",
        "invalid_contact": "❌ 请输入有效的邮箱或11位大陆手机号",
        "expander_activate_code": "🔑 输入激活码",
        "enter_activate_code_placeholder": "输入激活码",
        "no_license_hint": "💡 未激活，请扫码购买激活码",
        "valid_until": "⏰ 有效期至: {date}",
        "valid_until_date": "📅 有效期至: {date}",
        "permanent_valid": "♾️ 永久有效",
        "account_manager": "🔐 账户管理 / Account",
        "contact_service": "📱 联系客服 / Contact",
        "main_title": "📊 8D 报告智能生成助手",
        "main_subtitle": "AI 驱动的纠正预防措施报告",
        "progress_phases": [
            {"icon": "📝", "text": "正在整理您的输入信息...", "sub": "产品：{product}"},
            {"icon": "🤔", "text": "正在理解问题背景...", "sub": "运用 5W2H 方法分析"},
            {"icon": "📋", "text": "正在生成 D1 团队组建...", "sub": "确定责任人及时间节点"},
            {"icon": "📋", "text": "正在生成 D2 问题描述...", "sub": "详细记录不良现象"},
            {"icon": "🛡️", "text": "正在生成 D3 临时措施...", "sub": "遏制问题扩散"},
            {"icon": "🔬", "text": "正在分析根本原因 (4M1E)...", "sub": "人、机、料、法、环逐一排查"},
            {"icon": "🔍", "text": "正在进行 5-Why 追问...", "sub": "追溯至根本原因"},
            {"icon": "💡", "text": "正在制定 D5 永久措施...", "sub": "根本性解决方案"},
            {"icon": "✅", "text": "正在生成 D6 实施计划...", "sub": "验证措施有效性"},
            {"icon": "📊", "text": "正在生成 D7 预防措施...", "sub": "防止问题复发"},
            {"icon": "🏆", "text": "正在生成 D8 总结表彰...", "sub": "固化经验，分享成果"},
            {"icon": "✨", "text": "正在优化报告格式...", "sub": "确保专业美观"},
        ],         "input_header": "📝 输入基本信息",
        "card_product": "📋 产品信息",
        "card_problem": "⚠️ 问题描述",
        "card_details": "🔍 事件详情",
        "product_name": "产品型号 / 名称", "customer": "客户名称",
        "problem_desc": "不良现象描述",
        "problem_placeholder": "请使用 5W2H 方法描述问题",
        "occur_date": "发现日期", "defect_qty": "不良数量", "severity": "严重程度",
        "severity_low": "低", "severity_medium": "中", "severity_high": "高", "severity_critical": "危急",
        "industry_std": "适用标准", "team_members": "团队成员（可选）",
        "team_placeholder": "例：张明 (组长), 李华 (工程)",
        "generate_btn": "🚀 自动生成 8D 报告", 
        "generating": "8D 报告智能生成中，请稍候...",
        "preview_header": "📄 报告预览", "download_btn": "📥 导出 Word 报告",
        "download_ppt": "📊 导出 PPT 报告",
        "ppt_title": "8D 纠正措施报告",
        "export_disabled": "🔒 激活正式版后可导出 Word / PPT",
        "no_desc": "❌ 请输入不良现象描述",
        "no_license_error": "❌ 未激活，请购买激活码", "api_error": "❌ 服务异常",
        "success": "✅ 报告生成完成！", "report_complete": "报告生成完成！",
        "beautifying": "正在美化格式...", "word_title": "8D 问题纠正与预防措施报告",
        "system_error": "❌ 系统错误，请稍后重试",
        "history_header": "📋 生成历史",
        "load_report": "加载此报告",
        "delete_report": "删除",
        "no_history": "暂无历史记录",
        "edit_mode": "✏️ 编辑模式",
        "save_edit": "💾 保存修改",
        "edit_placeholder": "在此编辑报告内容...",
        "history_loaded": "已从历史记录加载",
    },
    "en": {
        "lang_label": "Language", "lang_zh": "中文", "lang_en": "English",
        "system_status": "System Status", "pro_version": "✅ Pro Version", "free_version": "⚠️ Not Activated",
        "license_valid_until": "📅 Valid until {exp}", "trial_used": "📊 Used {used} / {total}",
        "no_license": "⚠️ Not activated, please purchase activation code", "activate_title": "🔑 License / Renew",
        "activate_code_hint": "Activation Code", "activate_btn": "Activate",
        "activate_success": "✅ Activated successfully", "activate_fail": "❌ Invalid code",
        "invalid_activate_code": "Please enter a valid activation code",
        "license_expired": "❌ License expired",
        "login_required": "🔒 Please login", "logout": "Logout",
        "login_header": "👤 User Login",
        "username_label": "Email or Phone",
        "username_placeholder": "e.g., name@example.com or 13812345678",
        "login_register_btn": "🔓 Login / Register",
        "enter_username_error": "Please enter email or phone number",
        "invalid_email": "❌ Invalid email format, e.g., name@example.com",
        "invalid_phone": "❌ Invalid phone number, enter 11-digit mainland China number",
        "invalid_contact": "❌ Please enter a valid email or 11-digit phone number",
        "expander_activate_code": "🔑 Enter Activation Code",
        "enter_activate_code_placeholder": "Enter activation code",
        "no_license_hint": "💡 Not activated, please scan QR code to purchase activation code",
        "valid_until": "⏰ Valid until: {date}",
        "valid_until_date": "📅 Valid until: {date}",
        "permanent_valid": "♾️ Permanent",
        "account_manager": "🔐 Account Manager",
        "contact_service": "📱 Contact Service",

        "main_title": "📊 8D Report Generator",
        "main_subtitle": "AI-Powered Corrective Action Reports",
        "progress_phases": [
            {"icon": "📝", "text": "Organizing your input...", "sub": "Product: {product}"},
            {"icon": "🤔", "text": "Analyzing context...", "sub": "Using 5W2H method"},
            {"icon": "📋", "text": "Generating D1 Team...", "sub": "Defining responsibilities"},
            {"icon": "📋", "text": "Generating D2 Description...", "sub": "Recording defect details"},
            {"icon": "🛡️", "text": "Generating D3 Containment...", "sub": "Preventing spread"},
            {"icon": "🔬", "text": "Analyzing root cause (4M1E)...", "sub": "Checking all factors"},
            {"icon": "🔍", "text": "Performing 5-Why analysis...", "sub": "Finding root cause"},
            {"icon": "💡", "text": "Developing D5 Actions...", "sub": "Long-term solutions"},
            {"icon": "✅", "text": "Generating D6 Implementation...", "sub": "Verifying effectiveness"},
            {"icon": "📊", "text": "Generating D7 Prevention...", "sub": "Preventing recurrence"},
            {"icon": "🏆", "text": "Generating D8 Closure...", "sub": "Documenting lessons"},
            {"icon": "✨", "text": "Formatting report...", "sub": "Professional output"},
        ],         "input_header": "📝 Input Information",
        "card_product": "📋 Product Info",
        "card_problem": "⚠️ Problem Description",
        "card_details": "🔍 Event Details",
        "product_name": "Product Name / Model", "customer": "Customer Name",
        "problem_desc": "Problem Description",
        "problem_placeholder": "Please use 5W2H method",
        "occur_date": "Occurrence Date", "defect_qty": "Defect Quantity", "severity": "Severity",
        "severity_low": "Low", "severity_medium": "Medium", "severity_high": "High", "severity_critical": "Critical",
        "industry_std": "Standard", "team_members": "Team Members (Optional)",
        "team_placeholder": "e.g., Zhang(Leader), Li(Eng)",
        "generate_btn": "🚀 Generate 8D Report",
        "generating": "Generating report, please wait...",
        "preview_header": "📄 Report Preview", "download_btn": "📥 Export Word",
        "download_ppt": "📊 Export PPT",
        "ppt_title": "8D Corrective Action Report",
        "export_disabled": "🔒 Activate to export Word / PPT",
        "no_desc": "❌ Please enter description",
        "no_license_error": "❌ Not activated, please purchase activation code", "api_error": "❌ Service error",
        "success": "✅ Report generated!", "report_complete": "Report generated!",
        "beautifying": "Formatting...", "word_title": "8D Corrective Action Report",
        "system_error": "❌ System error, please try again later",
        "history_header": "📋 Generation History",
        "load_report": "Load this report",
        "delete_report": "Delete",
        "no_history": "No history yet",
        "edit_mode": "✏️ Edit Mode",
        "save_edit": "💾 Save Changes",
        "edit_placeholder": "Edit report content here...",
        "history_loaded": "Loaded from history",
    }
}

# ==================== 系统提示词 ====================
SYSTEM_PROMPT = {
    "zh": (
        "你是一位拥有 20 年经验的汽车电子行业高级质量工程师，精通 IATF 16949 标准和 8D 问题解决方法。"
        "请根据用户输入撰写专业、逻辑严密的 8D 报告。\n\n"
        "【8D 报告结构要求】\n"
        "报告必须严格按照以下 8 个步骤的顺序输出，不可颠倒：\n"
        "D1：建立团队（成立问题解决小组，列出成员及职责）\n"
        "D2：问题描述（使用 5W2H 方法描述问题：What、Why、Who、When、Where、How、How many）\n"
        "D3：制定临时控制措施（ICA，围堵措施，防止问题扩大）\n"
        "D4：根本原因分析（见下方详细要求）\n"
        "D5：制定永久纠正措施（PCA，针对根本原因的根本解决方案）\n"
        "D6：贯彻永久纠正措施（实施计划、验证有效性）\n"
        "D7：预防措施（防止类似问题在其他产品/流程中复发）\n"
        "D8：表彰小组（总结、表彰团队成员贡献）\n\n"
        "【D4 根本原因分析要求】\n"
        "根本原因分析必须包含两部分：产生原因 和 流出原因。\n\n"
        "一、产生原因分析（为什么会产生缺陷）：\n"
        "使用 4M1E 分析法（人、机、料、法、环）逐项确认，使用确定句而非疑问句：\n"
        "✅ 正常项：明确说明\"经检查，XX 符合标准，排除为根本原因\"\n"
        "❌ 异常项：明确说明\"经检查，XX 存在问题：[具体问题]\"\n"
        "❌ 不要使用\"是否\"、\"有没有\"等疑问句\n\n"
        "从异常项开始，使用 5-Why 分析法：\n"
        "连续追问\"为什么\"，至少追问 3-5 层，直到找到根本原因\n"
        "每层回答要具体，不能笼统\n\n"
        "二、流出原因分析（为什么缺陷没有被发现，流向了客户）：\n"
        "分析检验/拦截环节为什么会失效，同样使用 5-Why 分析法：\n"
        "Why1：为什么该缺陷在 XX 检验环节没有被发现？\n"
        "Why2：为什么检验标准/方法/频次存在漏洞？\n"
        "Why3：为什么检验人员没有执行到位？\n"
        "Why4：为什么检验流程设计不完善？\n"
        "Why5：为什么管理层没有重视检验环节？\n\n"
        "输出格式示例（注意换行）：\n"
        "【D4 根本原因分析】\n\n"
        "=== 产生原因分析 ===\n\n"
        "【4M1E 分析】\n\n"
        "人：经检查，操作员持证上岗 → 排除\n\n"
        "机：经检查，设备参数偏移 0.05mm → 异常项 ⚠️\n\n"
        "料：经检查，原材料合格 → 排除\n\n"
        "法：经检查，作业指导书过期 → 异常项 ⚠️\n\n"
        "环：经检查，环境符合要求 → 排除\n\n"
        "【5-Why 分析（产生原因）】\n\n"
        "Why1：为什么设备参数偏移？→ 传感器校准过期\n\n"
        "Why2：为什么校准过期？→ 年度校准计划未执行\n\n"
        "Why3：为什么计划未执行？→ 维护人员不足 ← 根本原因\n\n"
        "=== 流出原因分析 ===\n\n"
        "【检验环节失效分析】\n\n"
        "检验点：出货检验 OQC\n\n"
        "Why1：为什么偏移参数的产品流出了？→ OQC 检验标准未包含该参数\n\n"
        "Why2：为什么标准未包含？→ 控制计划未更新该参数\n\n"
        "Why3：为什么控制计划未更新？→ 工程变更流程缺失 ← 根本原因\n\n"
        "【真实性约束与置信度标注】\n"
        "你无法获知现场实测数据，严禁编造具体数值（温度、尺寸、电流、时长、批次量等）。\n"
        "每条根因与结论必须标注置信度：\n"
        "🟢 高置信度 = 来自用户输入中明确给出的事实\n"
        "🟡 中置信度 = 基于行业经验的合理推断\n"
        "🔴 低置信度 = AI 推测，需现场验证\n"
        "凡无法确认的数据，使用占位标记：[待现场确认]、[参数待实测]、[数据待补充]\n"
        "示例：机：回流焊峰值温度偏低 🟡（推测实际约 235°C，[参数待实测]）\n\n"
        "【其他要求】\n"
        "语气专业客观\n"
        "措施使用 [责任人 | 时间 | 状态] 格式\n"
        "报告正文不使用 Markdown 标记（末尾不再附信息完整性地图）\n"
        "直接输出 D1-D8 报告正文，末尾附信息完整性地图"
    ),
    
    "en": (
        "You are a Senior Quality Engineer with 20 years experience in automotive electronics, "
        "proficient in IATF 16949 and 8D methodology. Please write a professional 8D report based on user input.\n\n"
        "【8D Report Structure Requirements】\n"
        "The report MUST follow these 8 steps in strict order, do not swap them:\n"
        "D1: Establish Team (form the problem solving team, list members and roles)\n"
        "D2: Describe the Problem (use 5W2H: What, Why, Who, When, Where, How, How many)\n"
        "D3: Develop Interim Containment Plan (ICA, containment actions to prevent spread)\n"
        "D4: Root Cause Analysis (see detailed requirements below)\n"
        "D5: Develop Permanent Corrective Actions (PCA, solutions addressing root cause)\n"
        "D6: Implement and Validate Corrective Actions (implementation plan, verify effectiveness)\n"
        "D7: Preventive Measures (prevent recurrence in similar products/processes)\n"
        "D8: Recognize Team and Individual Contributions (conclude and recognize team)\n\n"
        "【D4 Root Cause Analysis Requirements】\n"
        "Root cause analysis MUST include two parts: Occurrence Cause and Escape Cause.\n\n"
        "Part 1 - Occurrence Cause (Why did the defect occur?):\n"
        "Use 4M1E analysis (Man, Machine, Material, Method, Environment) with declarative sentences:\n"
        "✅ Normal: 'Verified, XX meets standard, excluded as root cause'\n"
        "❌ Abnormal: 'Verified, XX has issue: [specific problem]'\n"
        "Then use 5-Why analysis from abnormal items, minimum 3-5 levels.\n\n"
        "Part 2 - Escape Cause (Why did the defect escape to customer?):\n"
        "Analyze why inspection/containment failed, use 5-Why:\n"
        "Why1: Why wasn't this defect caught at XX inspection?\n"
        "Why2: Why is there a gap in inspection standard/method/frequency?\n"
        "Why3: Why wasn't the inspector performing correctly?\n"
        "Why4: Why is the inspection process flawed?\n"
        "Why5: Why didn't management prioritize this?\n\n"
        "Format example:\n"
        "【D4 Root Cause Analysis】\n\n"
        "=== Occurrence Cause ===\n\n"
        "【4M1E Analysis】\n"
        "Man: Verified... Excluded\n\n"
        "【5-Why Analysis (Occurrence)】\n"
        "Why1: ...\n\n"
        "=== Escape Cause ===\n\n"
        "【Inspection Failure Analysis】\n"
        "Why1: ...\n\n"
        "【Truthfulness Constraints & Confidence Labeling】\n"
        "You do NOT have access to on-site measured data. Never fabricate specific values "
        "(temperature, dimension, current, duration, batch size, etc).\n"
        "Label confidence for every root cause and conclusion:\n"
        "🟢 High = fact explicitly given by the user\n"
        "🟡 Medium = reasonable inference from industry experience\n"
        "🔴 Low = AI speculation, needs on-site verification\n"
        "For any unverified data, use placeholders: [To be confirmed on-site], [Params to be measured], [Data to be filled]\n"
        "Example: Machine: reflow peak temp slightly low 🟡 (est. ~235°C, [Params to be measured])\n\n"
        "【Other Requirements】\n"
        "Professional tone\n"
        "Use [Owner|Date|Status] format for actions\n"
        "No Markdown in report body (table allowed in the map)\n"
        "Output D1-D8 directly. No Information Completeness Map."
    )
}

# ==================== 行业专属逻辑 ====================
# 每个行业带中英文标签与专属分析指引，用户选择后注入 prompt
INDUSTRIES = [
    {
        "zh": "通用制造", "en": "General Manufacturing",
        "zh_guide": "按通用 ISO 9001 质量管理原则分析，覆盖人/机/料/法/环全要素，无特殊行业合规要求。",
        "en_guide": "Follow general ISO 9001 principles covering Man/Machine/Material/Method/Environment, no special industry compliance.",
    },
    {
        "zh": "汽车电子 (IATF 16949)", "en": "Automotive (IATF 16949)",
        "zh_guide": "汽车行业专属：评估停线风险与召回影响；重视供应商变更管控（4M 变更）；围堵须覆盖在途品与客户端库存；永久措施纳入 SPC 监控，关键特性过程能力 Cpk≥1.33。",
        "en_guide": "Automotive specifics: assess line-down & recall risk; emphasize supplier change control (4M); containment covers in-transit & customer inventory; permanent actions under SPC with Cpk≥1.33.",
    },
    {
        "zh": "半导体 / 芯片", "en": "Semiconductor / Chip",
        "zh_guide": "半导体专属：Wafer Lot→Die→Package 逐级追溯；D4 走 FA 失效分析流程（X-Ray/SAM 非破坏检查 → 切片/SEM/EDX 破坏分析）；回流焊温度曲线（TAL/峰值温度）必查；MSA 测量系统分析 GR&R<10%。",
        "en_guide": "Semiconductor specifics: trace Wafer Lot→Die→Package; D4 FA flow (X-Ray/SAM non-destructive → cross-section/SEM/EDX); reflow profile (TAL/peak) required; MSA GR&R<10%.",
    },
    {
        "zh": "PCB / 电子制造", "en": "PCB / Electronics",
        "zh_guide": "PCB 专属：阻抗失效用 TDR 定位 + 切片测线宽/线距/介质厚度；分层/起泡用 288℃/10s 热应力试验；电镀查孔铜厚度分布；过程能力线宽/孔铜 Cpk≥1.33、阻抗 Cpk≥1.67；回流焊温度曲线与 AOI/X-Ray 关联分析。",
        "en_guide": "PCB specifics: TDR + cross-section for impedance; 288℃/10s thermal stress for delamination; plating thickness distribution; Cpk≥1.33 (width/hole), ≥1.67 (impedance); reflow profile linked to AOI/X-Ray.",
    },
    {
        "zh": "医疗器械 (ISO 13485)", "en": "Medical (ISO 13485)",
        "zh_guide": "医疗器械专属：强调法规合规与可追溯性（批次→患者）；根因须关联风险管理（ISO 14971）；措施考虑临床影响与上市后监督（PMS）。",
        "en_guide": "Medical specifics: regulatory compliance & traceability (lot→patient); link root cause to risk management (ISO 14971); consider clinical impact & post-market surveillance.",
    },
    {
        "zh": "航空航天 (AS9100)", "en": "Aerospace (AS9100)",
        "zh_guide": "航空航天专属：适航合规与安全性为首要；采用 FRACAS 故障报告、分析与纠正系统；措施需首件检验（FAI）与过程确认。",
        "en_guide": "Aerospace specifics: airworthiness & safety first; use FRACAS; actions require first-article inspection (FAI) & process qualification.",
    },
]

# ==================== 初始化配置 ====================
try:
    API_KEY = st.secrets["DEEPSEEK_API_KEY"]
    BASE_URL = st.secrets["DEEPSEEK_BASE_URL"]
except Exception:
    API_KEY = ""
    BASE_URL = "https://api.deepseek.com"

try:
    supabase = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
except Exception:
    supabase = None

# ==================== 核心功能函数 ====================
def get_user_license(user_id):
    return get_cached_license(user_id)

def create_free_license(user_id):
    if not supabase:
        return None
    try:
        r = supabase.table("licenses").insert({
            "user_id": user_id,
            "plan_type": "free",
            "trial_used": 0,
            "trial_limit": 0
        }).execute()
        return r.data[0] if r.data else None
    except Exception:
        return None

def can_generate_report(user_id):
    lic = get_user_license(user_id)
    if not lic:
        return False
    plan = lic.get('plan_type', 'free')
    if plan == 'free':
        return False  # 未激活
    # trial / pro / enterprise 统一只检查有效期
    expire = lic.get('license_expire')
    if expire:
        try:
            return datetime.now() < datetime.fromisoformat(expire)
        except Exception:
            return True
    # 没有设置有效期的正式版，允许使用
    if plan in ['pro', 'enterprise']:
        return True
    return False

def inc_trial_used(user_id):
    if not supabase:
        return
    try:
        lic = get_cached_license(user_id)
        if lic:
            new_count = lic.get('trial_used', 0) + 1
            supabase.table("licenses").update({"trial_used": new_count}).eq("user_id", user_id).execute()
            supabase.table("usage_logs").insert({
                "user_id": user_id,
                "action": "generate_report",
                "created_at": datetime.now().isoformat()
            }).execute()
            clear_license_cache(user_id)
    except Exception as e:
        logging.error(f"更新试用次数失败：{e}")

def activate_license_code(user_id, code):
    if not supabase:
        return False, "系统错误"
    try:
        code_clean = code.strip().upper()
        r = supabase.table("activation_codes").select("*").eq("code", code_clean).execute()
        if not r.data:
            return False, "无效的激活码"
        ac = r.data[0]
        if ac.get('is_used'):
            return False, "激活码已被使用"
        if ac.get('expire_date'):
            if datetime.now().date() > datetime.fromisoformat(ac['expire_date']).date():
                return False, "激活码已过期"
        # 原子抢码：仅当 is_used=False 时才更新，防止并发激活
        claim = supabase.table("activation_codes").update({
            "is_used": True,
            "used_by": user_id,
            "used_at": datetime.now().isoformat()
        }).eq("code", code_clean).eq("is_used", False).execute()
        # 如果没有行被更新，说明已被别人抢走
        if not claim.data or len(claim.data) == 0:
            return False, "激活码已被使用"
        duration = ac.get('duration_days') or 365
        exp_date = (datetime.now() + timedelta(days=duration)).isoformat()
        supabase.table("licenses").upsert({
            "user_id": user_id,
            "plan_type": ac.get('plan_type', 'pro'),
            "license_expire": exp_date,
            "trial_used": 0,
            "trial_limit": 0
        }, on_conflict="user_id").execute()
        clear_license_cache(user_id)
        formatted_date = exp_date[:10] if len(exp_date) >= 10 else exp_date
        return True, f"激活成功！有效期至 {formatted_date}"
    except Exception as e:
        logging.error(f"激活失败：{e}")
        return False, f"激活失败：{str(e)}"

def activate_trial_code(user_id, code):
    """激活试用码：0.99元获得2次试用"""
    if not supabase:
        return False, "系统错误"
    try:
        code_clean = code.strip().upper()
        r = supabase.table("trial_codes").select("*").eq("code", code_clean).execute()
        if not r.data:
            return False, "无效的试用码"
        tc = r.data[0]
        if tc.get('is_used'):
            return False, "试用码已被使用"
        # 原子抢码：仅当 is_used=False 时才更新，防止并发激活
        claim = supabase.table("trial_codes").update({
            "is_used": True,
            "used_by": user_id,
            "used_at": datetime.now().isoformat()
        }).eq("code", code_clean).eq("is_used", False).execute()
        # 如果没有行被更新，说明已被别人抢走
        if not claim.data or len(claim.data) == 0:
            return False, "试用码已被使用"
        # 给用户2次试用机会
        supabase.table("licenses").update({
            "trial_limit": 2,
            "trial_used": 0
        }).eq("user_id", user_id).execute()
        clear_license_cache(user_id)
        return True, "✅ 试用码激活成功！获得 2 次试用机会"
    except Exception as e:
        logging.error(f"试用码激活失败：{e}")
        return False, f"激活失败：{str(e)}"

def clean_format(text):
    if not text:
        return ""
    # 统一换行符：DeepSeek 输出偶发 \r\n，避免 setext 下划线正则因 \r 而漏判
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = text.replace("**", "").replace("#", "")
    for i in range(1, 9):
        text = re.sub(rf'(D{i}[:：])\s*\n+\s*', rf'\1 ', text)
    text = re.sub(r'([人机料法环]：)', r'\n\1', text)
    text = re.sub(r'(→ 排除|→ 异常项[^，]*？)', r'\1\n', text)
    text = re.sub(r'(Why\d+：)', r'\n\1', text)
    text = re.sub(r'(→ [^\n]+)(?=Why\d+：|$)', r'\1\n', text)
    text = re.sub(r'(为什么\d+：)', r'\n\1', text)
    text = re.sub(r'(→ [^\n]+)(?=为什么\d+：|$)', r'\1\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    # ★ 强力去 setext 下划线：
    # 1) 独占一行、整行只由 = 或 - 构成（含 :, 空格）→ 当作下划线，整行删掉
    # 2) 同一行的 === 装饰：`=== 标题 ===` → 去掉两侧 `===`，留下标题文本
    # 3) 同一行尾部/头部的 `===` 或 `---` 装饰（成对）→ 去掉
    # 这一步必须放在 D-section/Why 段落重排之后，避免误伤前面的换行规则
    text = re.sub(r'(?m)^\s*[=\-]{3,}[ \t:=]*\s*$', '', text)
    text = re.sub(r'(?m)^[ \t]*={2,}[ \t]*([^\n]+?)[ \t]*={2,}[ \t]*$', r'\1', text)
    text = re.sub(r'(?m)^[ \t]*-{2,}[ \t]*([^\n]+?)[ \t]*-{2,}[ \t]*$', r'\1', text)
    return re.sub(r'\n{3,}', '\n\n', text).strip()

def _ncols(row):
    """返回该行按 | 拆出的单元格数量（含空单元格），用于判断是否像表格行。"""
    s = row.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|'):
        s = s[:-1]
    return len(s.split('|'))


def _split_row(row):
    """把 markdown 表格的一行按 | 拆成单元格，去掉首尾多余的 |。"""
    s = row.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|'):
        s = s[:-1]
    return [c.strip() for c in s.split('|')]


def _is_table_start(lines, i):
    """判断 lines[i] 是否是一段表格的起始行。

    兼容两种情况：① 标准 markdown 表格（下一行是 |---|---| 分隔行）；
    ② 模型省略分隔行，仅靠连续的 | 分隔行构成表格。
    """
    if '|' not in lines[i] or _ncols(lines[i]) < 2:
        return False
    if i + 1 >= len(lines):
        return False
    nxt = lines[i + 1].strip().strip('|').strip()
    sep = bool(nxt) and '-' in nxt and all(c in '-:| ' for c in nxt)
    if sep:
        return True
    # 无分隔行：要求下一行同样是 >=2 列的 | 行（排除标题类行），才判定为表格
    if lines[i + 1].lstrip().startswith('【') or lines[i + 1].lstrip().startswith('#'):
        return False
    return '|' in lines[i + 1] and _ncols(lines[i + 1]) >= 2


def _esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _strip_map_title(sec):
    """去掉信息完整性地图原括号标题行，避免与卡片标题重复。"""
    first = sec.split('\n', 1)
    if re.search(r'信息完整性地图|Information Completeness Map', first[0]):
        return first[1].strip() if len(first) > 1 else ''
    return sec


def _render_body_md(body):
    """正文转「原生 markdown」：连续的 | 分隔行转成标准 markdown 表格（自动补分隔行），
    其余文本原样保留。返回原生 markdown 字符串，交给 st.markdown 渲染——
    不再手写 <table> HTML，从根本上杜绝标签泄漏为纯文本的问题。"""
    lines = body.split('\n')
    out = []
    i = 0
    n = len(lines)
    while i < n:
        if _is_table_start(lines, i):
            header = _split_row(lines[i])
            rows = []
            j = i + 1
            # 跳过模型可能写出的分隔行 |---|---|
            if j < n:
                s = lines[j].strip().strip('|').strip()
                if '-' in s and all(c in '-:| ' for c in s):
                    j += 1
            while j < n and lines[j].strip() and _ncols(lines[j]) >= 2:
                rows.append(_split_row(lines[j]))
                j += 1
            # 表格前补一个空行，确保 Streamlit 的 markdown 引擎能识别为表格
            if out and out[-1].strip():
                out.append('')
            # 输出标准 markdown 表格（Streamlit 原生支持，需带分隔行）
            out.append('| ' + ' | '.join(header) + ' |')
            out.append('| ' + ' | '.join('---' for _ in header) + ' |')
            for r in rows:
                out.append('| ' + ' | '.join(r) + ' |')
            out.append('')  # 表格后补一个空行，避免与下文粘连
            i = j
        else:
            out.append(lines[i])
            i += 1
    return '\n'.join(out)


def render_d_sections(content):
    """将报告内容按 D1-D8 拆分，渲染为带彩色边框的章节"""
    if not content:
        return
    # 清理 markdown 标记
    clean = content.replace("**", "").replace("#", "")
    # 按 D1-D8 拆分（不再单独处理信息完整性地图，AI 已不再生成）
    sections = re.split(r'\n(?=D[1-8][:：])', clean)
    d_found = False
    for sec in sections:
        sec = sec.strip()
        if not sec:
            continue
        # 跳过信息完整性地图残留（旧报告可能还有）
        if re.match(r'^[【\[]?信息完整性地图[】\]]?', sec) or re.match(r'^[【\[]?Information Completeness Map[】\]]?', sec):
            continue
        # 提取 D 编号
        m = re.match(r'(D[1-8])[:：]\s*(.*)', sec)
        if m:
            d_num = m.group(1)  # "D1"
            d_title = m.group(2).split('\n')[0]  # 第一行标题
            d_body = sec[len(m.group(0)):].strip()  # 剩余内容
            if not d_title:
                d_title = d_num
            css_class = f"d-section d-section-{d_num.lower()}"
            title_safe = _esc(d_title)
            body_md = _render_body_md(d_body)
            st.markdown(
                f'<div class="{css_class}">'
                f'<div class="d-section-title" style="color:#ffffff !important;">{d_num}：{title_safe}</div>'
                f'<div class="d-section-body" style="color:#ffffff !important;">{body_md}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
            d_found = True
        elif not d_found:
            # D1 之前的前言内容或无 D 章节的内容：包一层带白色内联样式的 div，保证可见
            st.markdown(
                f'<div class="report-preamble" style="color:#ffffff !important; background:transparent;">{_render_body_md(sec)}</div>',
                unsafe_allow_html=True
            )

def render_d0_card(product_name, customer, problem_desc, defect_qty, severity, industry_std, team_members, lang):
    """D0 前置自诊断卡片：根据表单输入实时计算问题分类、数据完整度、复杂度、紧急程度与围堵建议"""
    if lang == "zh":
        labels = ["问题分类", "数据完整度", "复杂度", "紧急程度"]
        tip_header = "💡 围堵建议"
        cls_map = {
            "外观": "外观缺陷", "功能": "功能失效", "尺寸": "尺寸/公差",
            "性能": "性能衰减", "装配": "装配不良", "物料": "物料/批次",
            "软件": "软件/逻辑", "其他": "其他",
        }
        comp_map = {"低": "低", "中": "中", "高": "高"}
        sev_map = {
            "critical": ("🔴", "紧急", "#dc2626"),
            "high": ("🟠", "高", "#ea580c"),
            "medium": ("🟡", "中", "#ca8a04"),
            "low": ("🟢", "低", "#16a34a"),
        }
    else:
        labels = ["Type", "Data Completeness", "Complexity", "Urgency"]
        tip_header = "💡 Containment Advice"
        cls_map = {
            "外观": "Appearance", "功能": "Function", "尺寸": "Dimension",
            "性能": "Performance", "装配": "Assembly", "物料": "Material",
            "软件": "Software", "其他": "Other",
        }
        comp_map = {"低": "Low", "中": "Medium", "高": "High"}
        sev_map = {
            "critical": ("🔴", "Critical", "#dc2626"),
            "high": ("🟠", "High", "#ea580c"),
            "medium": ("🟡", "Medium", "#ca8a04"),
            "low": ("🟢", "Low", "#16a34a"),
        }

    # ---- 问题分类（关键词匹配）----
    kw = {
        "外观": ["划伤", "刮伤", "异色", "变色", "脏污", "毛刺", "起泡", "开裂", "裂纹", "破损", "变形", "生锈", "缺料", "烧焦"],
        "功能": ["失效", "故障", "不工作", "无法", "异常", "死机", "黑屏", "失灵", "通讯", "通信", "误判", "短路", "开路", "击穿"],
        "尺寸": ["尺寸", "公差", "超差", "平面度", "厚度", "长度", "孔径", "偏移"],
        "性能": ["性能", "参数", "指标", "衰减", "漂移", "温升", "噪声", "阻抗"],
        "装配": ["装配", "错位", "漏装", "错装", "松动", "间隙", "干涉"],
        "物料": ["物料", "批次", "混料", "供应商", "原料", "变更"],
        "软件": ["软件", "程序", "代码", "逻辑", "算法", "固件"],
    }
    desc = problem_desc or ""
    ptype = "其他"
    for k, words in kw.items():
        if any(w in desc for w in words):
            ptype = k
            break
    ptype_label = cls_map.get(ptype, ptype)

    # ---- 数据完整度评分（满分 100）----
    score = 0
    if product_name: score += 15
    if customer: score += 10
    score += min(30, len(desc) // 8)  # 描述越长越完整，上限 30
    score += 10  # 发生日期（默认有值）
    if defect_qty and defect_qty > 0: score += 10
    score += 10  # 严重程度（必选）
    score += 5   # 行业标准
    if team_members: score += 10
    score = max(0, min(100, score))
    # 数据完整度颜色：≥80 绿，50-79 黄，<50 红
    bar_color = "#16a34a" if score >= 80 else "#ca8a04" if score >= 50 else "#dc2626"

    # ---- 复杂度评级 ----
    if defect_qty and defect_qty > 1000 or len(desc) > 200:
        complexity = "高"
    elif (defect_qty and defect_qty > 100) or len(desc) > 80:
        complexity = "中"
    else:
        complexity = "低"
    complexity_label = comp_map.get(complexity, complexity)

    # ---- 紧急程度（来自严重程度）----
    sev_key = None
    sl = TEXT[lang]
    if severity == sl["severity_critical"]: sev_key = "critical"
    elif severity == sl["severity_high"]: sev_key = "high"
    elif severity == sl["severity_medium"]: sev_key = "medium"
    else: sev_key = "low"
    icon, urg_label, urg_color = sev_map[sev_key]

    # ---- 围堵建议 ----
    if sev_key in ("critical", "high"):
        tip = ("问题紧急，建议立即启动围堵措施（ICA），优先拦截在途品与客户端库存。"
               if lang == "zh" else
               "Urgent — recommend launching interim containment (ICA) immediately, prioritize intercepting in-transit & customer inventory.")
    else:
        tip = ("建议评估影响范围后启动围堵措施，避免问题扩大。"
               if lang == "zh" else
               "Recommend containment after assessing impact scope to prevent spread.")

    html = f'''
    <div class="d0-card">
      <div class="d0-card-header">🔍 D0 前置诊断 / Pre-check</div>
      <div class="d0-row"><span>{labels[0]}</span><b>{ptype_label}</b></div>
      <div class="d0-row"><span>{labels[1]}</span>
        <div class="d0-bar"><div class="d0-bar-fill" style="width:{score}%;background:{bar_color}"></div></div>
        <b style="color:{bar_color}">{score}%</b>
      </div>
      <div class="d0-row"><span>{labels[2]}</span><b>{complexity_label}</b></div>
      <div class="d0-row"><span>{labels[3]}</span><b style="color:{urg_color}">{icon} {urg_label}</b></div>
      <div class="d0-tip"><b>{tip_header}：</b>{tip}</div>
    </div>
    '''
    st.markdown(html, unsafe_allow_html=True)


def _add_body_to_doc(doc, body):
    """把正文写入 Word：含 markdown 表格则生成真实表格，否则按段写入。"""
    lines = body.split('\n')
    text_buf = []
    i = 0
    n = len(lines)

    def flush_text():
        if text_buf:
            doc.add_paragraph(''.join(text_buf).strip())
            text_buf.clear()

    while i < n:
        if _is_table_start(lines, i):
            flush_text()
            header = _split_row(lines[i])
            j = i + 1
            if j < n:
                s = lines[j].strip().strip('|').strip()
                if '-' in s and all(c in '-:| ' for c in s):
                    j += 1
            rows = []
            while j < n and lines[j].strip() and _ncols(lines[j]) >= 2:
                rows.append(_split_row(lines[j]))
                j += 1
            table = doc.add_table(rows=1, cols=max(1, len(header)))
            try:
                table.style = 'Light Grid Accent 1'
            except Exception:
                pass
            hdr = table.rows[0].cells
            for k, h in enumerate(header):
                hdr[k].text = h
            for r in rows:
                cells = table.add_row().cells
                for k, c in enumerate(r):
                    cells[k].text = c
            i = j
        else:
            text_buf.append(lines[i] + '\n')
            i += 1
    flush_text()


def export_to_word(content, product_name, lang):
    doc = Document()
    # 页面背景改为浅蓝色（默认是白色）
    sectPr = doc.sections[0]._sectPr
    background = OxmlElement('w:background')
    background.set(qn('w:color'), 'E6EEF8')
    sectPr.append(background)
    if lang == "zh":
        doc.styles['Normal'].font.name = '宋体'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    else:
        doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10.5)
    title = doc.add_heading(TEXT[lang]["word_title"], 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = doc.add_paragraph()
    info.add_run(f"Product: {product_name}").bold = True
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    clean = content.replace("**", "").replace("#", "")
    sections = re.split(r'\n(?=D[1-8][:：])', clean)
    for i, sec in enumerate(sections):
        if not sec.strip():
            continue
        # 跳过残留的旧信息完整性地图章节
        if re.match(r'^[【\[]?信息完整性地图[】\]]?', sec.strip()) or re.match(r'^[【\[]?Information Completeness Map[】\]]?', sec.strip()):
            continue
        # 普通 D 章节
        lines = sec.strip().split('\n', 1)
        p_title = doc.add_paragraph()
        runner = p_title.add_run(lines[0].strip())
        runner.bold = True
        runner.font.size = Pt(14)
        runner.font.color.rgb = RGBColor(30, 58, 138)
        if len(lines) > 1 and lines[1].strip():
            _add_body_to_doc(doc, lines[1].strip())
        if i < len(sections) - 1:
            p_line = doc.add_paragraph()
            p_line.paragraph_format.space_before = Pt(12)
            p_line.paragraph_format.space_after = Pt(12)
            p = p_line._element
            pPr = p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '8')
            pBdr.append(bottom)
            pPr.append(pBdr)
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def _ppt_parse_blocks(body):
    """把章节正文解析成结构化块：('heading',..) / ('para',..) / ('bullet',..) / ('table',(header,rows))"""
    if not body:
        return []
    lines = body.split('\n')
    blocks = []
    i, n = 0, len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        # 表格
        if _is_table_start(lines, i):
            header = _split_row(lines[i])
            rows = []
            j = i + 1
            if j < n:
                ss = lines[j].strip().strip('|').strip()
                if '-' in ss and all(c in '-:| ' for c in ss):
                    j += 1
            while j < n and lines[j].strip() and _ncols(lines[j]) >= 2:
                rows.append(_split_row(lines[j]))
                j += 1
            blocks.append(('table', (header, rows)))
            i = j
            continue
        # 标题：=== xxx ===
        if s.startswith('===') and s.endswith('==='):
            blocks.append(('heading', s.strip('= ').strip()))
            i += 1
            continue
        # 小标题：以 ：结尾且较短
        if (s.endswith('：') or s.endswith(':')) and len(s) <= 28:
            blocks.append(('heading', s))
            i += 1
            continue
        # 项目符号
        if re.match(r'^([-\u2022\u25cf*]|\d+[.)]|[a-zA-Z][.)])\s+', s) or s.startswith('\u2192') or s.startswith('\u25b6') or s.startswith('\u2022'):
            txt = re.sub(r'^([-\u2022\u25cf*]|\d+[.)]|[a-zA-Z][.)]|\u2192|\u25b6)\s*', '', s).strip()
            blocks.append(('bullet', txt))
            i += 1
            continue
        blocks.append(('para', s))
        i += 1
    return blocks


def export_to_pptx(content, product_name, lang):
    """将报告导出为结构化的 PPT：标题页 + 每个 D 章节一页（含分色标题带、徽章、表格、自动分页）。"""
    from pptx import Presentation
    from pptx.util import Pt, Inches
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE

    D_COLORS = {
        "D1": RGBColor(0x25, 0x63, 0xEB),
        "D2": RGBColor(0x16, 0xA3, 0x4A),
        "D3": RGBColor(0xEA, 0x58, 0x0C),
        "D4": RGBColor(0xDB, 0x27, 0x77),
        "D5": RGBColor(0x93, 0x33, 0xEA),
        "D6": RGBColor(0x08, 0x91, 0xB2),
        "D7": RGBColor(0xCA, 0x8A, 0x04),
        "D8": RGBColor(0x64, 0x74, 0x8B),
        "MAP": RGBColor(0x47, 0x55, 0x69),
    }
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    DARK = RGBColor(0x33, 0x33, 0x33)
    GREY = RGBColor(0x94, 0xA3, 0xB8)
    NAVY = RGBColor(0x1E, 0x3A, 0x8A)
    LIGHT_BLUE = RGBColor(0xE6, 0xEE, 0xF8)   # 浅蓝（与网页卡片同色）

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    SW = float(prs.slide_width) / 914400.0
    SH = float(prs.slide_height) / 914400.0

    # ---------- 标题页 ----------
    ts = prs.slides.add_slide(prs.slide_layouts[6])
    ts.background.fill.solid()
    ts.background.fill.fore_color.rgb = NAVY
    bar = ts.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.1), prs.slide_width, Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(0x60, 0xA5, 0xFA)
    bar.line.fill.background()
    t = ts.shapes.add_textbox(Inches(1), Inches(2.1), SW - 2, Inches(1.2))
    tf = t.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = TEXT[lang]["ppt_title"]
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = WHITE
    sub = ts.shapes.add_textbox(Inches(1), Inches(4.0), SW - 2, Inches(0.8))
    stf = sub.text_frame
    stf.word_wrap = True
    sp = stf.paragraphs[0]
    sp.alignment = PP_ALIGN.CENTER
    sr = sp.add_run()
    sr.text = f"Product / 产品: {product_name}"
    sr.font.size = Pt(20)
    sr.font.color.rgb = RGBColor(0xBF, 0xDB, 0xFE)
    dt = ts.shapes.add_textbox(Inches(1), Inches(5.4), SW - 2, Inches(0.5))
    dtf = dt.text_frame
    dtf.paragraphs[0].alignment = PP_ALIGN.CENTER
    dr = dtf.paragraphs[0].add_run()
    dr.text = datetime.now().strftime("%Y-%m-%d")
    dr.font.size = Pt(14)
    dr.font.color.rgb = GREY

    # ---------- 章节拆分 ----------
    clean = content.replace("**", "").replace("#", "")
    sections = re.split(
        r'\n(?=D[1-8][:：])',
        clean
    )

    MARGIN = 0.6
    TOP = 1.2
    BOTTOM = SH - 0.6
    X = MARGIN
    W = SW - 2 * MARGIN

    def _make_slide(color, badge, title, continuation=False):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = LIGHT_BLUE
        # 顶部色带
        band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.95))
        band.fill.solid()
        band.fill.fore_color.rgb = color
        band.line.fill.background()
        # 徽章
        badge_shape = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.4), Inches(0.18), Inches(0.6), Inches(0.6))
        badge_shape.fill.solid()
        badge_shape.fill.fore_color.rgb = WHITE
        badge_shape.line.fill.background()
        bt = badge_shape.text_frame
        bt.word_wrap = False
        bp = bt.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = badge
        br.font.size = Pt(16)
        br.font.bold = True
        br.font.color.rgb = color
        # 标题
        tt = s.shapes.add_textbox(Inches(1.15), Inches(0.18), SW - 1.5, Inches(0.6))
        ttf = tt.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.LEFT
        tr = tp.add_run()
        tr.text = title + ("（续）" if continuation else "")
        tr.font.size = Pt(22)
        tr.font.bold = True
        tr.font.color.rgb = WHITE
        # 页脚
        ft = s.shapes.add_textbox(Inches(0.5), SH - 0.45, SW - 1, Inches(0.35))
        ftf = ft.text_frame
        ftf.paragraphs[0].alignment = PP_ALIGN.LEFT
        fr = ftf.paragraphs[0].add_run()
        fr.text = f"{product_name}  ·  8D Report  ·  {datetime.now().strftime('%Y-%m-%d')}"
        fr.font.size = Pt(9)
        fr.font.color.rgb = GREY
        return s

    def _est(text, pt, w_in):
        cw = pt * 0.55 / 72.0
        nlines = max(1, -(-len(text) // max(1, int(w_in / cw))))
        return nlines * (pt * 1.25 / 72.0) + 0.08

    for sec in sections:
        sec = sec.strip()
        if not sec:
            continue
        lines = sec.split('\n')
        head = lines[0].strip()
        body = '\n'.join(lines[1:]).strip()
        m = re.match(r'D(\d)[:：]\s*(.*)', head)
        # 跳过残留的旧信息完整性地图章节
        if re.match(r'^[【\[]?信息完整性地图[】\]]?', head) or re.match(r'^[【\[]?Information Completeness Map[】\]]?', head):
            continue
        if m:
            dnum = "D" + m.group(1)
            dtitle = m.group(2).strip() or dnum
            color = D_COLORS.get(dnum, D_COLORS["D8"])
            badge = m.group(1)
        else:
            dnum = ""
            dtitle = head[:40]
            color = D_COLORS["D8"]
            badge = "•"

        blocks = _ppt_parse_blocks(body)
        if not blocks:
            blocks = [('para', body if body else head)]

        y = TOP
        s = _make_slide(color, badge, dtitle)
        for kind, payload in blocks:
            if kind == 'table':
                header, rows = payload
                nrows = len(rows) + 1
                ncols = max(len(header), 1)
                h = nrows * 0.32 + 0.1
                if y + h > BOTTOM and y > TOP:
                    s = _make_slide(color, badge, dtitle, True)
                    y = TOP
                tbl_shape = s.shapes.add_table(nrows, ncols, Inches(X), Inches(y), Inches(W), Inches(h))
                tbl = tbl_shape.table
                for c in range(ncols):
                    tbl.columns[c].width = Inches(W / ncols)
                for c in range(ncols):
                    cell = tbl.cell(0, c)
                    cell.text = header[c] if c < len(header) else ""
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = color
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    pr = cell.text_frame.paragraphs[0]
                    pr.alignment = PP_ALIGN.CENTER
                    run = pr.runs[0]
                    run.font.size = Pt(11)
                    run.font.bold = True
                    run.font.color.rgb = WHITE
                for ri, row in enumerate(rows, start=1):
                    for c in range(ncols):
                        cell = tbl.cell(ri, c)
                        cell.text = row[c] if c < len(row) else ""
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = WHITE if ri % 2 else RGBColor(0xF1, 0xF5, 0xF9)
                        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                        pr = cell.text_frame.paragraphs[0]
                        run = pr.runs[0]
                        run.font.size = Pt(10)
                        run.font.color.rgb = DARK
                y = y + h + 0.15
            else:
                text = payload
                pt = 18 if kind == 'heading' else 14
                h = _est(text, pt, W)
                if y + h > BOTTOM and y > TOP:
                    s = _make_slide(color, badge, dtitle, True)
                    y = TOP
                tb = s.shapes.add_textbox(Inches(X), Inches(y), Inches(W), Inches(h))
                tf = tb.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = ("•  " + text) if kind == 'bullet' else text
                for run in p.runs:
                    run.font.size = Pt(pt)
                    run.font.color.rgb = DARK
                if kind == 'heading':
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = color
                    p.space_after = Pt(4)
                y = y + h + (0.12 if kind in ('para', 'bullet') else 0.18)

    bio = BytesIO()
    prs.save(bio)
    return bio.getvalue()


# ==================== 历史记录功能 ====================
def save_report_history(user_id, product_name, customer, problem_desc, report_content, lang):
    """保存报告到历史记录"""
    if not supabase or not user_id:
        return
    try:
        supabase.table("reports").insert({
            "user_id": user_id,
            "product_name": product_name or "",
            "customer": customer or "",
            "problem_desc": (problem_desc or "")[:500],
            "report_content": report_content,
            "lang": lang,
            "created_at": datetime.now().isoformat()
        }).execute()
    except Exception as e:
        logging.warning(f"保存历史记录失败：{e}")

def load_report_history(user_id, limit=10):
    """加载用户历史记录"""
    if not supabase or not user_id:
        return []
    try:
        r = supabase.table("reports").select("*").eq("user_id", user_id).order("created_at", desc=True).limit(limit).execute()
        return r.data or []
    except Exception as e:
        logging.warning(f"加载历史记录失败：{e}")
        return []

def delete_report_history(report_id):
    """删除单条历史记录"""
    if not supabase:
        return False
    try:
        supabase.table("reports").delete().eq("id", report_id).execute()
        return True
    except Exception as e:
        logging.warning(f"删除历史记录失败：{e}")
        return False

# ==================== 会话状态初始化 ====================
if "lang" not in st.session_state:
    st.session_state.lang = "zh"
if "current_result" not in st.session_state:
    st.session_state.current_result = ""
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "registration_attempted" not in st.session_state:
    st.session_state.registration_attempted = False

T = TEXT[st.session_state.lang]

# ==================== 侧边栏（登录和用户管理） ====================
def render_sidebar():
    """渲染侧边栏 - 登录、用户信息、语言切换、激活码等"""
    T = TEXT[st.session_state.lang]
    
    with st.sidebar:
        # ==================== 语言切换 ====================
        st.markdown("### 🌐 语言 / Language")
        lang_option = st.selectbox(
            "选择语言 / Select Language",
            ["中文", "English"],
            index=0 if st.session_state.lang == "zh" else 1,
            key="sidebar_lang_select",
            label_visibility="collapsed"
        )
        new_lang = "zh" if lang_option == "中文" else "en"
        if new_lang != st.session_state.lang:
            st.session_state.lang = new_lang
            st.rerun()
        
        st.markdown(f"### {T['account_manager']}")
                
        # ==================== 登录 / 注册区域 ====================
        user_id = st.session_state.get("user_id")
        
        if not user_id:
            user_input = st.text_input(
                T["username_label"],
                key="sidebar_user_input",
                placeholder=T["username_placeholder"]
            )
            
            # ========== 格式校验函数 ==========
            def validate_contact(contact):
                """校验邮箱或手机号格式"""
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                phone_pattern = r'^1[3-9]\d{9}$'
                
                if re.match(email_pattern, contact):
                    return True, "email"
                elif re.match(phone_pattern, contact):
                    return True, "phone"
                else:
                    return False, None
            
            st.caption("💡 首次输入将自动创建账号，无需单独注册")
            
            if st.button(T["login_register_btn"], use_container_width=True, key="sidebar_login_btn"):
                if not user_input:
                    st.error(T["enter_username_error"])
                    st.stop()

                # ========== 方案1：会话级注册限制 ==========
                if st.session_state.registration_attempted:
                    st.error("⚠️ 当前会话已注册过，请勿重复操作" if st.session_state.lang == "zh" else "⚠️ Already registered in this session")
                    st.stop()

                # 检查是否为已注册的老用户
                existing_user = False
                if supabase:
                    try:
                        existing = supabase.table("licenses").select("user_id").eq("user_id", user_input).execute()
                        existing_user = existing.data is not None and len(existing.data) > 0
                    except Exception:
                        st.error(T["system_error"])
                        st.stop()

                # 校验格式
                is_valid, contact_type = validate_contact(user_input)

                # 老用户不受格式限制，直接放行
                if not existing_user and not is_valid:
                    if "@" in user_input:
                        st.error(T["invalid_email"])
                    elif user_input.startswith("1") and len(user_input) == 11:
                        st.error(T["invalid_phone"])
                    else:
                        st.error(T["invalid_contact"])
                    st.stop()

                # ========== 方案2：相似账号检测（仅新用户） ==========
                if not existing_user and supabase and user_input.isdigit() and len(user_input) >= 10:
                    try:
                        prefix_len = len(user_input) - 2
                        prefix = user_input[:prefix_len]
                        similar = supabase.table("licenses").select("user_id").like("user_id", prefix + "%").limit(5).execute()
                        if similar.data and len(similar.data) > 0:
                            st.error(
                                "⚠️ 检测到可疑注册行为，已被拒绝。请联系客服。"
                                if st.session_state.lang == "zh" else
                                "⚠️ Suspicious registration detected. Please contact support."
                            )
                            st.stop()
                    except Exception as e:
                        logging.warning(f"相似账号检测失败：{e}")

                # 新用户注册（格式验证通过 且 无历史记录）
                if not existing_user and supabase:
                    try:
                        supabase.table("licenses").insert({
                            "user_id": user_input,
                            "plan_type": "free",
                            "trial_used": 0,
                            "trial_limit": 0
                        }).execute()
                        st.session_state.registration_attempted = True  # 标记会话已注册
                    except Exception:
                        pass

                st.session_state.user_id = user_input
                st.rerun()
        
        # ==================== 已登录用户区域 ====================
        else:
            lic = get_user_license(user_id)
            
            st.markdown(f"**👤 {user_id[:30]}**")
            
            if lic:
                if lic.get('plan_type') == 'free':
                    st.warning(T["no_license"])
                    # 购买引导
                    st.markdown("---")
                    st.markdown("### 💰 购买正式版")
                    try:
                        st.image("paid.jpg", width=200)
                    except:
                        st.info("请上传 paid.jpg 到项目目录")
                    st.markdown("""
**版本与价格：**

| 版本 | 原价 | 优惠价 |
|------|------|--------|
| 月卡 | ~~¥29~~ | **¥9.9/月** |
| 年卡 | ~~¥99~~ | **¥39/年** |
| 5年卡 | ~~¥299~~ | **¥99/5年** |

**购买步骤：**
1. 截图上面的二维码
2. 微信扫码转账（选对应金额）
3. 转账后联系微信 **907749064** 获取激活码
4. 在下方"🔑 输入激活码"中输入激活码
                    """)
                else:
                    st.success(T["pro_version"])
                    if lic.get('license_expire'):
                        try:
                            exp_date = datetime.fromisoformat(lic['license_expire']).strftime('%Y-%m-%d')
                            st.caption(T["valid_until_date"].format(date=exp_date))
                        except:
                            pass
                    else:
                        st.caption(T["permanent_valid"])
            
                        
            with st.expander(T["expander_activate_code"], expanded=False):
                activate_code = st.text_input(
                    "输入激活码",
                    type="password",
                    key="sidebar_act_code",
                    placeholder="输入激活码"
                )
                if st.button("激活", key="sidebar_act_btn", use_container_width=True):
                    if activate_code and len(activate_code) >= 6:
                        code_upper = activate_code.strip().upper()
                        if code_upper.startswith("8DT1"):
                            success, msg = activate_trial_code(user_id, activate_code)
                        elif code_upper.startswith(("8D8P", "8D30", "8D8E")):
                            success, msg = activate_license_code(user_id, activate_code)
                        else:
                            success, msg = False, "请输入有效的激活码（8DT1 / 8D8P / 8D30 / 8D8E 开头）"
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.error("请输入有效的激活码")
            
            if st.button(T["logout"], key="sidebar_logout_btn", use_container_width=True):
                st.session_state.user_id = None
                st.session_state.current_result = ""
                get_cached_license.clear()
                st.rerun()
        
        # ==================== 历史记录 ====================
        if user_id:
            st.markdown(f"**{T['history_header']}**")
            history = load_report_history(user_id)
            if not history:
                st.caption(T["no_history"])
            else:
                for report in history[:5]:
                    with st.expander(f"{(report.get('product_name') or 'N/A')[:20]} | {report['created_at'][:10]}"):
                        st.caption((report.get('problem_desc') or '')[:80])
                        col_load, col_del = st.columns([3, 1])
                        with col_load:
                            if st.button(T["load_report"], key=f"load_{report['id']}", use_container_width=True):
                                st.session_state.current_result = report['report_content']
                                st.success(T["history_loaded"])
                                st.rerun()
                        with col_del:
                            if st.button("🗑️", key=f"del_{report['id']}", use_container_width=True):
                                if delete_report_history(report['id']):
                                    st.rerun()

        st.markdown("---")

        # ==================== 底部信息 ====================
        
        st.markdown(f"**{T['contact_service']}**")
        try:
            st.image("wechat_qrcode.jpg", width=180)
        except:
            st.info("微信二维码：907749064")
        st.caption("淘宝店铺: 效率工坊铺")
        st.caption("微信号Wechat: 907749064")
        st.caption("Email: 907749064@qq.com")
        st.markdown("---")
 
# ==================== 主页面 ====================
render_sidebar()

st.markdown(f'''
<div class="brand-header">
    <div class="brand-logo">8D</div>
    <div>
        <div class="brand-title">{T["main_title"]}</div>
        <div class="brand-subtitle">{T["main_subtitle"]}</div>
    </div>
</div>
''', unsafe_allow_html=True)

col_input, col_preview = st.columns([1, 1.2])

with col_input:
    st.markdown(f'<div class="panel-title">{T["input_header"]}</div>', unsafe_allow_html=True)

    # ★ 关键改动：原本三个独立的 st.container(border=True) 会留出 Streamlit 默认空隙。
    # 现在把整段塞进"一个"带边框容器，从根本上消除卡片之间的间距。
    # 卡头以横向分隔线 <div class="form-divider"> 视觉区分，不再占额外高度。
    with st.container(border=True):
        # ========== 第 1 段：产品 & 客户 ==========
        c1, c2 = st.columns(2)
        with c1:
            product_name = st.text_input(T["product_name"], placeholder="例：PCB-A123" if st.session_state.lang == "zh" else "e.g., PCB-A123")
        with c2:
            customer = st.text_input(T["customer"], placeholder="例：比亚迪汽车" if st.session_state.lang == "zh" else "e.g., BYD")

        # 分隔线
        st.markdown('<div class="form-divider"></div>', unsafe_allow_html=True)

        # ========== 第 2 段：问题描述 ==========
        problem_desc = st.text_area(T["problem_desc"], height=90, placeholder=T["problem_placeholder"])

        # 分隔线
        st.markdown('<div class="form-divider"></div>', unsafe_allow_html=True)

        # ========== 第 3 段：事件详情 ==========
        col1, col2, col3 = st.columns(3)
        with col1:
            occur_date = st.date_input(T["occur_date"], datetime.now())
        with col2:
            defect_qty = st.number_input(T["defect_qty"], min_value=1, value=1)
        with col3:
            severity = st.selectbox(
                T["severity"],
                [T["severity_low"], T["severity_medium"], T["severity_high"], T["severity_critical"]]
            )

        col4, col5 = st.columns(2)
        with col4:
            industry_opts = [x["zh"] if st.session_state.lang == "zh" else x["en"] for x in INDUSTRIES]
            industry_std = st.selectbox(
                T["industry_std"],
                industry_opts,
                index=1
            )
        with col5:
            team_members = st.text_input(T["team_members"], placeholder=T["team_placeholder"])

        # 解析选中的行业，提取专属指引
        sel_industry = next(
            (x for x in INDUSTRIES if x["zh"] == industry_std or x["en"] == industry_std),
            INDUSTRIES[1]
        )
        industry_guide = sel_industry[st.session_state.lang + "_guide"]

        # ========== D0 前置自诊断卡片 ==========
        render_d0_card(product_name, customer, problem_desc, defect_qty, severity, industry_std, team_members, st.session_state.lang)
    
        
    if st.button(T["generate_btn"], type="primary", use_container_width=True):
        if not st.session_state.get("user_id"):
            st.error(T["login_required"])
            st.stop()

        user_id = st.session_state.user_id
        if not can_generate_report(user_id):
            lic = get_user_license(user_id)
            if lic and lic['plan_type'] == 'free':
                st.error(T["no_license"])
            else:
                st.error(T["license_expired"])
            st.stop()

        if not problem_desc:
            st.error(T["no_desc"])
        else:
            with st.status(T["generating"], expanded=True) as status:
                try:
                    client = openai.OpenAI(api_key=API_KEY, base_url=BASE_URL)

                    if st.session_state.lang == "zh":
                        user_prompt = (
                            f"请根据以下信息生成 8D 报告："
                            f"产品：{product_name or '未提供'}, "
                            f"客户：{customer or '未提供'}, "
                            f"日期：{occur_date}, "
                            f"数量：{defect_qty}, "
                            f"严重程度：{severity}, "
                            f"标准：{industry_std}, "
                            f"团队：{team_members or '未提供'}\n\n"
                            f"问题描述：{problem_desc}\n\n"
                            f"【行业专属要求】\n{industry_guide}"
                        )
                    else:
                        user_prompt = (
                            f"Generate 8D report based on:\n"
                            f"Product: {product_name or 'N/A'}\n"
                            f"Customer: {customer or 'N/A'}\n"
                            f"Date: {occur_date}\n"
                            f"Quantity: {defect_qty}\n"
                            f"Severity: {severity}\n"
                            f"Standard: {industry_std}\n"
                            f"Team: {team_members or 'N/A'}\n\n"
                            f"Problem Description: {problem_desc}\n\n"
                            f"[Industry-specific Requirements]\n{industry_guide}"
                        )

                    response = client.chat.completions.create(
                        model="deepseek-chat",
                        messages=[
                            {"role": "system", "content": SYSTEM_PROMPT[st.session_state.lang]},
                            {"role": "user", "content": user_prompt}
                        ],
                        stream=True,
                        temperature=0.2,
                        max_tokens=4096
                    )

                    full_content = ""
                    progress_placeholder = st.empty()
                    stream_placeholder = st.empty()

                    for chunk in response:
                        delta = chunk.choices[0].delta.content
                        if delta:
                            full_content += delta
                            # 检测当前 D 步骤
                            current_d = 0
                            for i in range(1, 9):
                                if re.search(rf'D{i}[:：]', full_content):
                                    current_d = i
                            # 更新进度圆点
                            if current_d > 0:
                                dots_html = '<div class="progress-dots">'
                                for i in range(1, 9):
                                    if i < current_d:
                                        dots_html += '<div class="progress-dot dot-done">✓</div>'
                                    elif i == current_d:
                                        dots_html += f'<div class="progress-dot dot-active">D{i}</div>'
                                    else:
                                        dots_html += f'<div class="progress-dot dot-pending">D{i}</div>'
                                dots_html += '</div>'
                                progress_placeholder.markdown(dots_html, unsafe_allow_html=True)
                            stream_placeholder.markdown(full_content)

                    # 完成后全部绿色
                    dots_done = '<div class="progress-dots">'
                    for i in range(1, 9):
                        dots_done += '<div class="progress-dot dot-done">✓</div>'
                    dots_done += '</div>'
                    progress_placeholder.markdown(dots_done, unsafe_allow_html=True)

                    status.update(label="✅ " + T["success"], state="complete", expanded=False)

                    final_result = clean_format(full_content)
                    st.session_state.current_result = final_result
                    inc_trial_used(user_id)
                    save_report_history(user_id, product_name, customer, problem_desc, final_result, st.session_state.lang)

                except openai.APIConnectionError:
                    status.update(label="❌ 网络连接失败", state="error")
                    st.error("🌐 网络连接失败，请检查网络后重试")
                except openai.RateLimitError:
                    status.update(label="❌ API 频率超限", state="error")
                    st.error("⏱️ API 调用频率超限，请等待 30 秒后重试")
                except openai.AuthenticationError:
                    status.update(label="❌ API 密钥验证失败", state="error")
                    st.error("🔑 API 密钥验证失败，请联系管理员")
                except openai.APIError as e:
                    status.update(label="❌ 服务异常", state="error")
                    err_detail = str(e) if str(e) else "未知错误"
                    logging.error(f"APIError 详情：{e}", exc_info=True)
                    st.error(f"❌ 服务异常：{err_detail}")
                except Exception as e:
                    status.update(label="❌ 系统错误", state="error")
                    logging.error(f"生成报告未知错误：{e}", exc_info=True)
                    st.error(T["api_error"])

with col_preview:
    st.markdown(f'<div class="panel-title">{T["preview_header"]}</div>', unsafe_allow_html=True)
    if st.session_state.current_result:
        edit_mode = st.checkbox(T["edit_mode"], key="edit_mode_toggle")
        if edit_mode:
            edited = st.text_area(T["edit_placeholder"], value=st.session_state.current_result, height=600, key="edit_area")
            if st.button(T["save_edit"], use_container_width=True, key="save_edit_btn"):
                st.session_state.current_result = edited
                st.session_state.edit_mode_toggle = False
                st.success("✅ 修改已保存" if st.session_state.lang == "zh" else "✅ Changes saved")
                st.rerun()
        else:
            render_d_sections(st.session_state.current_result)
        
        st.markdown("---")
        
        # ========== 按钮行：复制 + 导出并排 ==========
        user_id = st.session_state.get("user_id")
        lic = get_user_license(user_id) if user_id else None
        
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            # ========== 一键复制按钮 ==========
            copy_b64 = base64.b64encode(st.session_state.current_result.encode('utf-8')).decode('ascii')
            copy_label = "📋 一键复制报告" if st.session_state.lang == "zh" else "📋 Copy Report"
            copied_label = "✅ 已复制到剪贴板" if st.session_state.lang == "zh" else "✅ Copied!"
            fail_label = "复制失败，请手动选择文本复制" if st.session_state.lang == "zh" else "Copy failed"

            copy_html = f'''
            <div style="width:100%;">
            <button id="copy-btn" style="
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                border: none;
                padding: 0.45rem 1rem;
                border-radius: 0.5rem;
                cursor: pointer;
                font-size: 0.9rem;
                width: 100%;
            " onclick="
                try {{
                    const b64 = '{copy_b64}';
                    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
                    const text = new TextDecoder('utf-8').decode(bytes);
                    const ta = document.createElement('textarea');
                    ta.value = text;
                    ta.style.position = 'fixed';
                    ta.style.opacity = '0';
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                    const btn = document.getElementById('copy-btn');
                    btn.textContent = '{copied_label}';
                    btn.style.background = 'linear-gradient(135deg, #11998e 0%, #38ef7d 100%)';
                    setTimeout(function() {{
                        btn.textContent = '{copy_label}';
                        btn.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
                    }}, 2000);
                }} catch(e) {{
                    alert('{fail_label}');
                }}
            ">{copy_label}</button>
            </div>
            '''
            components.html(copy_html, height=45)
        
        with btn_col2:
            if lic and lic['plan_type'] != 'free':
                ex_col1, ex_col2 = st.columns(2)
                with ex_col1:
                    word_data = export_to_word(
                        st.session_state.current_result,
                        product_name or "8D_Report",
                        st.session_state.lang
                    )
                    st.download_button(
                        label=T["download_btn"],
                        data=word_data,
                        file_name=f"8D_Report_{datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with ex_col2:
                    try:
                        ppt_data = export_to_pptx(
                            st.session_state.current_result,
                            product_name or "8D_Report",
                            st.session_state.lang
                        )
                        st.download_button(
                            label=T["download_ppt"],
                            data=ppt_data,
                            file_name=f"8D_Report_{datetime.now().strftime('%Y%m%d')}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                    except Exception:
                        st.warning("⚠️ PPT 导出不可用（依赖缺失）" if st.session_state.lang == "zh" else "⚠️ PPT export unavailable")
            else:
                st.info(T["export_disabled"])
    else:
        st.info("👈 输入问题描述后点击生成" if st.session_state.lang == "zh" else "👈 Enter description and click generate")
