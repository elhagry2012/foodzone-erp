# -*- coding: utf-8 -*-
"""
واجهة الويب لنظام الحاجري المالي | ELHAGRY Financial Systems (Streamlit) - تعمل محلياً وعلى Streamlit Cloud.

التشغيل محلياً:
    pip install -r requirements.txt
    streamlit run streamlit_app.py

تستخدم نفس المحرك المحاسبي (الحزمة elhagry_core) وقاعدة البيانات والصلاحيات التي تستخدمها واجهة سطح المكتب
(python foodzone_erp.py). التصميم متجاوب بالكامل: موبايل، تابلت، لابتوب.
"""
import base64
import html
import io
import os
import re
import sqlite3
import tempfile
import time
from datetime import date, timedelta

import pandas as pd
import streamlit as st
from PIL import Image

import elhagry_core as F

APP_TITLE = F.BRAND_FULL
SESSION_TIMEOUT_MIN = 30
# لوحة الألوان: أزرق نابض + فيروزي + أردوازي
BLUE, BLUE_DK, TEAL, CYAN, SLATE, NAVY = "#2563eb", "#1e3a8a", "#14b8a6", "#06b6d4", "#0f172a", "#0b1e3f"
ON_STREAMLIT_CLOUD = os.path.isdir("/mount/src") or bool(os.environ.get("STREAMLIT_SHARING_MODE"))


# ============================================================
# التهيئة والشعار
# ============================================================
@st.cache_resource(show_spinner="جاري تجهيز قاعدة البيانات...")
def boot():
    F.init_db()
    try:
        F.auto_backup_daily()
    except Exception:
        pass
    return True


@st.cache_resource
def logo_image():
    """شعار الحاجري الرسمي (elhagry.png) مقتطعاً من تصميمه، مع لون خلفيته."""
    path = F.find_logo_file()
    if not path:
        return None, NAVY
    try:
        return F.extract_logo_image(Image.open(path))
    except Exception:
        return None, NAVY


@st.cache_resource
def logo_data_uri(max_w=360):
    im, bg = logo_image()
    if im is None:
        return None, bg
    im = im.copy().convert("RGBA")
    im.thumbnail((max_w, max_w))
    try:  # إزالة لون خلفية الشعار (يصبح شفافاً بحواف ناعمة)
        br, bgg, bb = (int(bg[i:i + 2], 16) for i in (1, 3, 5))
        px = im.load()
        for y in range(im.height):
            for x in range(im.width):
                r, g, b, a = px[x, y]
                d = max(abs(r - br), abs(g - bgg), abs(b - bb))
                if d < 60:
                    px[x, y] = (r, g, b, int(a * max(0, d - 14) / 46))
    except (ValueError, TypeError):
        pass
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode(), bg


def page_icon():
    im, _ = logo_image()
    return im if im is not None else "📊"


st.set_page_config(page_title=APP_TITLE, page_icon=page_icon(), layout="wide", initial_sidebar_state="auto")

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
:root {{
  --blue:{BLUE}; --blue-dk:{BLUE_DK}; --teal:{TEAL}; --cyan:{CYAN}; --slate:{SLATE}; --navy:{NAVY};
  --ink:#0f172a; --muted:#5b6b82; --line:rgba(148,163,184,.28);
  --card-shadow:0 1px 2px rgba(15,23,42,.06), 0 14px 30px -16px rgba(15,23,42,.30);
  --lift-shadow:0 2px 4px rgba(15,23,42,.06), 0 26px 44px -20px rgba(37,99,235,.45);
  --grad:linear-gradient(135deg, var(--blue) 0%, #0ea5e9 55%, var(--teal) 100%);
}}
html, body, .stApp, [data-testid="stSidebar"], button, input, textarea, select, label, p, li, td, th {{
  font-family:'Cairo','Tahoma',sans-serif;
}}
.stApp {{
  direction:rtl;
  background:
    radial-gradient(1100px 520px at 105% -8%, rgba(37,99,235,.16) 0%, transparent 60%),
    radial-gradient(900px 480px at -8% 108%, rgba(20,184,166,.16) 0%, transparent 58%),
    #f1f5fb;
}}
.block-container {{ padding-top:1.1rem; max-width:1440px; }}
h1, h2, h3, h4, h5, h6, p, label, li, .stMarkdown, .stCaption, [data-testid="stMetricLabel"] {{ text-align:right; }}
h3, h4 {{ color:var(--blue-dk); font-weight:800; }}
input, textarea {{ text-align:right; direction:rtl; }}
input[type="number"] {{ direction:ltr; text-align:right; }}
[data-testid="stDataFrame"], [data-testid="stTable"] {{ direction:ltr; }}

/* ---------- الشريط الجانبي */
[data-testid="stSidebar"] {{
  background:linear-gradient(185deg, {NAVY} 0%, #0f1d3d 45%, #0b1220 100%);
  box-shadow:-10px 0 34px rgba(2,6,23,.35); border-left:1px solid rgba(255,255,255,.05);
}}
[data-testid="stSidebar"] > div {{ direction:rtl; text-align:right; }}
[data-testid="stSidebar"] * {{ color:#dbe4f3; }}
[data-testid="stSidebar"] [role="radiogroup"] {{ gap:3px; }}
[data-testid="stSidebar"] [role="radiogroup"] label {{
  width:100%; padding:8px 12px; border-radius:12px; margin:0; transition:all .18s ease; border:1px solid transparent;
}}
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]) {{ display:none; }}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {{
  background:rgba(56,189,248,.10); border-color:rgba(56,189,248,.25); transform:translateX(-3px);
}}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{
  background:var(--grad); box-shadow:0 10px 22px -8px rgba(37,99,235,.75), inset 0 1px 0 rgba(255,255,255,.25);
}}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) * {{ color:#fff !important; font-weight:700; }}
[data-testid="stSidebar"] .stButton > button {{
  background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.14); color:#e2e8f0 !important;
}}
[data-testid="stSidebar"] .stButton > button:hover {{ background:rgba(239,68,68,.18); border-color:rgba(239,68,68,.5); }}
.fz-side-brand {{
  text-align:center; border-radius:18px; padding:16px 10px 12px; margin:4px 0 12px;
  box-shadow:0 18px 34px -18px rgba(0,0,0,.8), inset 0 0 0 1px rgba(56,189,248,.18);
}}
.fz-side-brand img {{ width:62%; max-width:170px; filter:drop-shadow(0 10px 18px rgba(37,99,235,.35)); }}
.fz-side-brand .ar {{ color:#fff !important; font-weight:800; font-size:1.05rem; margin-top:4px; }}
.fz-side-brand .en {{ color:{CYAN} !important; letter-spacing:2px; font-size:.68rem; font-weight:700; }}
.fz-user-chip {{
  display:flex; align-items:center; gap:10px; background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.10);
  border-radius:14px; padding:8px 12px; margin:6px 0 10px;
}}
.fz-user-chip .av {{
  width:34px; height:34px; border-radius:50%; background:var(--grad); display:flex; align-items:center;
  justify-content:center; color:#fff !important; font-weight:800; flex:none;
}}
.fz-user-chip .n {{ font-weight:700; color:#fff !important; line-height:1.1; }}
.fz-user-chip .r {{ font-size:.75rem; color:#93c5fd !important; }}

/* ---------- رأس الصفحة */
.fz-header {{
  position:relative; overflow:hidden; display:flex; align-items:center; justify-content:space-between; gap:14px;
  background:linear-gradient(120deg, {BLUE_DK} 0%, {BLUE} 42%, #0891b2 78%, {TEAL} 100%);
  color:#fff; border-radius:20px; padding:18px 26px; margin-bottom:16px;
  box-shadow:0 22px 44px -22px rgba(37,99,235,.75), inset 0 1px 0 rgba(255,255,255,.28);
}}
.fz-header::before, .fz-header::after {{
  content:""; position:absolute; border-radius:50%; background:rgba(255,255,255,.10); pointer-events:none;
}}
.fz-header::before {{ width:240px; height:240px; left:-60px; top:-120px; }}
.fz-header::after {{ width:160px; height:160px; left:140px; bottom:-110px; background:rgba(255,255,255,.07); }}
.fz-header .t {{ font-size:1.45rem; font-weight:800; text-shadow:0 2px 10px rgba(0,0,0,.18); }}
.fz-header .b {{ font-size:.8rem; opacity:.9; letter-spacing:.5px; }}
.fz-header .u {{
  background:rgba(255,255,255,.16); border:1px solid rgba(255,255,255,.28); backdrop-filter:blur(6px);
  border-radius:999px; padding:6px 14px; font-size:.85rem; font-weight:700; white-space:nowrap; z-index:1;
}}

/* ---------- البطاقات (عمق ثلاثي الأبعاد) */
[data-testid="stMetric"], .fz-card, [data-testid="stExpander"], [data-testid="stForm"] {{
  background:#fff; border:1px solid var(--line); border-radius:18px; box-shadow:var(--card-shadow);
  transition:transform .18s ease, box-shadow .18s ease;
}}
[data-testid="stMetric"] {{ position:relative; overflow:hidden; padding:14px 18px 12px; }}
[data-testid="stMetric"]::before {{
  content:""; position:absolute; inset:0 0 auto 0; height:4px; background:var(--grad);
}}
[data-testid="stMetric"]:hover, .fz-card:hover {{ transform:translateY(-4px); box-shadow:var(--lift-shadow); }}
[data-testid="stMetricLabel"] p {{ color:var(--muted); font-weight:700; font-size:.85rem; }}
[data-testid="stMetricValue"] {{ font-size:1.4rem; font-weight:800; color:var(--ink); direction:ltr; text-align:right; }}
.fz-card {{ padding:16px 20px; margin-bottom:12px; }}
.fz-card h4 {{ margin:0 0 6px; }}
.fz-chip {{
  display:inline-block; background:linear-gradient(135deg,#eff6ff,#ecfeff); color:var(--blue-dk);
  border:1px solid #bfdbfe; border-radius:999px; padding:2px 10px; margin:2px; font-size:.78rem; font-weight:700;
}}
.fz-steps {{ display:flex; gap:10px; flex-wrap:wrap; }}
.fz-step {{
  flex:1 1 200px; background:#fff; border:1px solid var(--line); border-radius:16px; padding:12px 14px;
  box-shadow:var(--card-shadow); transition:transform .18s ease, box-shadow .18s ease;
}}
.fz-step:hover {{ transform:translateY(-3px); box-shadow:var(--lift-shadow); }}
.fz-step b {{
  display:inline-flex; width:26px; height:26px; border-radius:50%; background:var(--grad); color:#fff;
  align-items:center; justify-content:center; margin-left:6px; box-shadow:0 6px 14px -6px rgba(37,99,235,.8);
}}
[data-testid="stDataFrame"] {{
  background:#fff; border-radius:16px; overflow:hidden; border:1px solid var(--line); box-shadow:var(--card-shadow);
}}
[data-testid="stForm"] {{ padding:18px 20px; }}
[data-testid="stAlert"] {{ border-radius:14px; box-shadow:0 8px 20px -14px rgba(15,23,42,.35); }}

/* ---------- التبويبات */
[data-testid="stTabs"] [role="tablist"] {{
  gap:6px; background:rgba(255,255,255,.78); padding:6px; border-radius:16px; border:1px solid var(--line);
  box-shadow:inset 0 1px 2px rgba(15,23,42,.06), 0 10px 22px -18px rgba(15,23,42,.4); flex-wrap:wrap;
}}
[data-testid="stTab"] {{ border-radius:12px; padding:8px 16px !important; font-weight:700; height:auto; transition:all .16s; }}
[data-testid="stTab"]:hover {{ background:#eff6ff; }}
[data-testid="stTab"][aria-selected="true"] {{ background:var(--grad); box-shadow:0 10px 20px -10px rgba(37,99,235,.8); }}
[data-testid="stTab"][aria-selected="true"] p {{ color:#fff !important; }}

/* ---------- الأزرار والحقول */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {{
  border-radius:12px; font-weight:700; background:#fff; color:var(--blue-dk); border:1px solid rgba(37,99,235,.28);
  box-shadow:0 2px 0 rgba(15,23,42,.05), 0 10px 20px -14px rgba(15,23,42,.45); transition:all .16s ease;
}}
.stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {{
  transform:translateY(-2px); border-color:var(--blue); color:var(--blue);
  box-shadow:0 16px 26px -14px rgba(37,99,235,.6);
}}
.stButton > button[data-testid="stBaseButton-primary"],
.stFormSubmitButton > button[data-testid="stBaseButton-primaryFormSubmit"],
.stButton > button[kind="primary"] {{
  background:var(--grad) !important; color:#fff !important; border:0 !important;
  box-shadow:0 12px 26px -12px rgba(37,99,235,.85), inset 0 1px 0 rgba(255,255,255,.3);
}}
.stButton > button[data-testid="stBaseButton-primary"]:hover,
.stFormSubmitButton > button[data-testid="stBaseButton-primaryFormSubmit"]:hover {{
  filter:brightness(1.08); color:#fff !important; transform:translateY(-2px);
}}
.stButton > button[data-testid="stBaseButton-primary"] p,
.stFormSubmitButton > button[data-testid="stBaseButton-primaryFormSubmit"] p {{ color:#fff !important; }}
.stButton > button:disabled {{ opacity:.55; transform:none; }}
[data-testid="stTextInputRootElement"], [data-testid="stNumberInputContainer"], [data-testid="stDateInputField"],
[data-testid="stSelectbox"] [role="group"], [data-testid="stMultiSelect"] [role="group"], [data-testid="stTextArea"] textarea {{
  border-radius:12px !important; background:#f8fafc !important; border:1px solid #c7d2e3 !important;
  transition:border-color .15s, box-shadow .15s;
}}
[data-testid="stTextInputRootElement"] input, [data-testid="stNumberInputContainer"] input {{ background:transparent !important; }}
[data-testid="stTextInputRootElement"]:focus-within, [data-testid="stNumberInputContainer"]:focus-within,
[data-testid="stDateInputField"]:focus-within, [data-testid="stSelectbox"] [role="group"]:focus-within {{
  border-color:var(--teal) !important; box-shadow:0 0 0 3px rgba(20,184,166,.22); background:#fff !important;
}}
[data-testid="stFileUploaderDropzone"] {{
  border:2px dashed #93c5fd; border-radius:18px; background:linear-gradient(135deg,#eff6ff 0%,#f0fdfa 100%);
  transition:all .18s ease;
}}
[data-testid="stFileUploaderDropzone"]:hover {{ border-color:var(--teal); box-shadow:var(--lift-shadow); }}

/* ---------- شاشة الدخول */
.fz-login-panel {{
  border-radius:26px; padding:42px 22px 26px; text-align:center; color:#fff; position:relative; overflow:hidden;
  box-shadow:0 34px 60px -28px rgba(11,30,63,.9), inset 0 0 0 1px rgba(56,189,248,.20);
}}
.fz-login-panel img {{ width:74%; max-width:300px; filter:drop-shadow(0 18px 28px rgba(37,99,235,.35)); }}
.fz-login-panel .en {{ color:{CYAN}; letter-spacing:3px; font-size:.78rem; font-weight:700; margin-top:10px; }}
.fz-login-panel .ar {{ color:#fff; font-size:1.25rem; font-weight:800; margin-top:6px; }}
.fz-login-title {{ font-size:2rem; font-weight:800; color:var(--ink); margin-bottom:0; }}
.fz-login-sub {{ color:var(--muted); margin-bottom:10px; }}

/* ---------- قائمة الدخل */
.fz-income {{ width:100%; border-collapse:separate; border-spacing:0; direction:rtl; font-size:.95rem; background:#fff;
  border-radius:18px; overflow:hidden; box-shadow:var(--card-shadow); border:1px solid var(--line); }}
.fz-income td {{ padding:7px 14px; border-bottom:1px solid #eef2f7; }}
.fz-income td.v {{ text-align:left; direction:ltr; font-variant-numeric:tabular-nums; white-space:nowrap; }}
.fz-income tr:hover td {{ background:#f8fbff; }}
.fz-income tr.h td {{ font-weight:800; background:linear-gradient(90deg,#eff6ff,#ecfeff); color:var(--blue-dk); }}
.fz-income tr.sec td {{ font-weight:800; color:var(--blue); padding-top:14px; }}
/* ---------- قائمة التنقل المجمعة */
.fz-nav-h {{ color:#7b8db0 !important; font-size:.72rem; font-weight:800; letter-spacing:.5px; margin:12px 6px 4px; }}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] button {{
  justify-content:flex-start; text-align:right; background:transparent !important; border:1px solid transparent !important;
  box-shadow:none !important; color:#dbe4f3 !important; padding:6px 12px; min-height:38px; border-radius:12px;
}}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] button p {{ color:#dbe4f3 !important; font-weight:600; text-align:right; }}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] button > div {{ width:100%; justify-content:flex-start !important; }}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] {{ margin-top:-12px; }}
[data-testid="stSidebar"] .fz-nav-h {{ margin-bottom:12px; }}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] button:hover {{
  background:rgba(56,189,248,.10) !important; border-color:rgba(56,189,248,.25) !important; transform:translateX(-3px);
}}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] button[data-testid="stBaseButton-primary"] {{
  background:var(--grad) !important; box-shadow:0 10px 22px -8px rgba(37,99,235,.75) !important;
}}
[data-testid="stSidebar"] [class*="st-key-nav_btn_"] button[data-testid="stBaseButton-primary"] p {{ color:#fff !important; font-weight:800; }}
[data-testid="stSidebar"] .st-key-nav_btn_import button {{ border:1px solid rgba(34,211,238,.45) !important; }}
[data-testid="stSidebar"] .st-key-nav_btn_import button p {{ color:#67e8f9 !important; }}
[data-testid="stSidebar"] .st-key-nav_btn_import button::after {{
  content:"Excel / CSV"; margin-right:auto; font-size:.62rem; font-weight:800; color:#0f172a; background:#67e8f9;
  border-radius:999px; padding:1px 8px; white-space:nowrap;
}}
[data-testid="stSidebar"] .st-key-nav_btn_discounts button p {{ color:#5eead4 !important; }}
[data-testid="stSidebar"] .st-key-logout_btn button:hover {{ background:rgba(239,68,68,.18) !important; border-color:rgba(239,68,68,.5) !important; }}

/* ---------- محدد الأقسام داخل الصفحات (أزرار حبوب) */
[data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] {{ gap:8px; flex-wrap:wrap; }}
[data-testid="stMain"] [data-testid="stRadioOption"] {{
  background:#fff; border:1px solid var(--line); border-radius:12px; padding:7px 16px; box-shadow:var(--card-shadow);
  transition:all .16s ease; cursor:pointer;
}}
[data-testid="stMain"] [data-testid="stRadioOption"]:hover {{ transform:translateY(-2px); border-color:var(--blue); }}
[data-testid="stMain"] [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]) {{ display:none; }}
[data-testid="stMain"] [data-testid="stRadioOption"]:has(input:checked) {{
  background:var(--grad); border-color:transparent; box-shadow:0 10px 22px -10px rgba(37,99,235,.8);
}}
[data-testid="stMain"] [data-testid="stRadioOption"]:has(input:checked) p {{ color:#fff !important; font-weight:800; }}
/* ============================================================
   التصميم المتجاوب: لابتوب (> 1024px) / تابلت (641–1024px) / موبايل (≤ 640px)
   ============================================================ */
html, body {{ overflow-x:hidden; }}
/* الشريط الجانبي المطوي: يختفي تماماً (مع اتجاه RTL كان يظهر منه شريط رفيع فوق المحتوى على الموبايل) */
[data-testid="stSidebar"][data-collapsed="true"], [data-testid="stSidebar"][aria-expanded="false"] {{
  visibility:hidden; min-width:0 !important; width:0 !important; overflow:hidden;
}}
[data-testid="stMain"] {{ overflow-x:hidden; }}
.block-container {{ padding-left:1.4rem; padding-right:1.4rem; }}
/* الجداول: تمرير أفقي سلس بدلاً من انضغاط النصوص */
[data-testid="stDataFrame"], [data-testid="stTable"], [data-testid="stDataFrameResizable"] {{
  max-width:100%; overflow-x:auto; -webkit-overflow-scrolling:touch; overscroll-behavior-x:contain;
}}
[data-testid="stTable"] table {{ min-width:560px; }}
.fz-scroll {{ width:100%; overflow-x:auto; -webkit-overflow-scrolling:touch; border-radius:18px; }}
.fz-scroll > table {{ min-width:520px; }}
.fz-income td:first-child {{ white-space:normal; word-break:break-word; }}
/* الأعمدة والبطاقات: التفاف بدل التداخل */
[data-testid="stMain"] [data-testid="stHorizontalBlock"] {{ flex-wrap:wrap; row-gap:.6rem; }}
[data-testid="stMain"] [data-testid="stColumn"], [data-testid="stMain"] [data-testid="column"] {{ min-width:0; }}
[data-testid="stMetricValue"] > div {{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
[data-testid="stMetricLabel"] p {{ white-space:normal; line-height:1.3; }}
.fz-step {{ min-width:0; }}
img, svg, canvas {{ max-width:100%; }}
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {{ min-height:42px; }}

/* ---------- بطاقات الجداول: مخفية على الشاشات الواسعة، وتحل محل الجدول على الموبايل */
.fz-cards {{ display:none; }}
.fz-card-row {{ background:#fff; border:1px solid var(--line); border-radius:14px; padding:10px 12px; margin-bottom:8px;
  box-shadow:0 6px 16px -12px rgba(15,23,42,.35); }}
.fz-card-row .t {{ font-weight:800; color:var(--blue-dk); margin-bottom:6px; white-space:normal; word-break:break-word; }}
.fz-kv {{ display:flex; justify-content:space-between; gap:10px; padding:3px 0; border-top:1px dashed #eef2f7; }}
.fz-kv .k {{ color:var(--muted); font-size:.82rem; flex:0 0 auto; max-width:48%; }}
.fz-kv .v {{ font-weight:700; text-align:left; white-space:normal; word-break:break-word; }}
.fz-kv .v.n {{ direction:ltr; font-variant-numeric:tabular-nums; }}
.fz-kv.hl {{ background:#fef08a; border-radius:8px; padding:3px 6px; }}
.fz-cards-more {{ color:var(--muted); font-size:.8rem; text-align:center; padding:6px; }}

/* ---------- مساحة العمل: تبويبات متجاورة (تقرير | كشف | فاتورة | سداد) */
.st-key-ws_bar {{ justify-content:space-between; flex-wrap:wrap; row-gap:2px; }}
.st-key-ws_tab [data-testid="stTab"] p {{ max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.st-key-cv_pay_summary [data-testid="stMetric"] {{ min-width:150px; }}
.st-key-cv_pay_summary > div {{ flex:1 1 150px; min-width:0; }}

/* ---------- تابلت */
@media (max-width: 1024px) {{
  .st-key-ws_tab [data-testid="stTab"] p {{ max-width:210px; font-size:.88rem; }}
  .block-container {{ padding:.9rem 1rem 3rem; max-width:100%; }}
  [data-testid="stMain"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"],
  [data-testid="stMain"] [data-testid="stHorizontalBlock"] > [data-testid="column"] {{
    flex:1 1 calc(50% - 1rem) !important; min-width:calc(50% - 1rem) !important; width:auto !important;
  }}
  [data-testid="stMetricValue"] {{ font-size:1.2rem; }}
  .fz-header {{ padding:16px 18px; border-radius:18px; }}
  .fz-header .t {{ font-size:1.25rem; }}
  [data-testid="stTab"] {{ padding:7px 12px !important; }}
  .fz-login-panel {{ padding:30px 18px 20px; }}
}}

/* ---------- موبايل */
@media (max-width: 640px) {{
  .block-container {{ padding:.6rem .7rem 3rem; }}
  [data-testid="stMain"] [data-testid="stHorizontalBlock"] {{ flex-direction:column; gap:.5rem; }}
  [data-testid="stMain"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"],
  [data-testid="stMain"] [data-testid="stHorizontalBlock"] > [data-testid="column"] {{
    flex:1 1 100% !important; min-width:100% !important; width:100% !important;
  }}
  .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {{ width:100%; }}
  .fz-header {{ flex-direction:column; align-items:flex-end; padding:14px 16px; gap:8px; }}
  .fz-header .t {{ font-size:1.1rem; }}
  .fz-header .u {{ white-space:normal; font-size:.78rem; }}
  .fz-header::before, .fz-header::after {{ display:none; }}
  [data-testid="stMetric"] {{ padding:10px 14px 8px; }}
  [data-testid="stMetricValue"] {{ font-size:1.1rem; }}
  [data-testid="stMetric"]:hover, .fz-card:hover, .fz-step:hover {{ transform:none; }}
  [data-testid="stTabs"] [role="tablist"] {{ flex-wrap:nowrap; overflow-x:auto; -webkit-overflow-scrolling:touch; }}
  [data-testid="stTab"] {{ flex:none; white-space:nowrap; padding:6px 10px !important; }}
  [data-testid="stForm"] {{ padding:12px 12px; }}
  /* تبويبات مساحة العمل على الموبايل: شريط واحد يُمرَّر أفقياً باللمس، وأزرار التبويب لا تتمدد لعرض الشاشة */
  .st-key-ws_tab [role="tablist"] {{ flex-wrap:nowrap; overflow-x:auto; scrollbar-width:none; }}
  .st-key-ws_tab [role="tablist"]::-webkit-scrollbar {{ display:none; }}
  .st-key-ws_tab [data-testid="stTab"] {{ min-height:42px; }}
  .st-key-ws_tab [data-testid="stTab"] p {{ max-width:170px; font-size:.82rem; }}
  .st-key-ws_bar .stButton > button, [class*="st-key-ws_x_"] .stButton > button,
  [class*="st-key-inv_inline_x_"] .stButton > button {{ width:auto; min-height:40px; }}
  .st-key-cv_pay_summary {{ flex-direction:column !important; }}
  .st-key-cv_pay_summary > div {{ width:100% !important; }}
  .fz-card {{ padding:12px 14px; }}
  .fz-step {{ flex:1 1 100%; }}
  .fz-income {{ font-size:.86rem; }}
  .fz-income td {{ padding:6px 8px; }}
  input, textarea, select {{ font-size:16px !important; }}   /* يمنع التكبير التلقائي في iOS */
  /* الجداول للعرض فقط تتحول لبطاقات: لا تمرير أفقي */
  [class*="st-key-fzt_c_"] [data-testid="stDataFrame"],
  [class*="st-key-fzt_c_"] [data-testid="stElementContainer"]:has([data-testid="stDataFrame"]) {{ display:none; }}
  [class*="st-key-fzt_c_"] .fz-cards {{ display:block; }}
  .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {{ min-height:46px; }}
  [data-testid="stDataEditor"], [data-testid="stDataFrame"] {{ font-size:.9rem; }}
  .fz-login-panel {{ padding:24px 14px 16px; border-radius:20px; }}
  .fz-login-panel img {{ width:56%; }}
  .fz-login-title {{ font-size:1.5rem; }}
  [data-testid="stMain"] [data-testid="stRadioOption"] {{ padding:6px 12px; }}
  h3 {{ font-size:1.15rem; }} h4 {{ font-size:1.02rem; }}
}}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ============================================================
# أدوات مساعدة
# ============================================================
def U():
    return st.session_state.get("user")


def can(key):
    u = U()
    return bool(u) and F.user_can(u, key)


def can_fin(role):
    u = U()
    return bool(u) and F.user_can_fin(u, role)


def flash(msg, kind="success"):
    st.session_state.setdefault("_flash", []).append((kind, msg))


def show_flash():
    for kind, msg in st.session_state.pop("_flash", []):
        getattr(st, kind)(msg)


def run_safe(fn, *args, **kwargs):
    """ينفذ عملية ويعرض أخطاء المستخدم كتنبيه مفهوم، وأي عطل غير متوقع كرسالة عامة دون كشف التفاصيل."""
    try:
        return True, fn(*args, **kwargs)
    except F.UserError as e:
        st.error(str(e))
    except sqlite3.OperationalError as e:
        st.error("قاعدة البيانات مشغولة أو مقفلة مؤقتاً، أعد المحاولة بعد لحظات." if "locked" in str(e)
                 else f"خطأ في قاعدة البيانات: {e}")
    except Exception as e:  # noqa: BLE001
        st.error(f"حدث خطأ غير متوقع: {e}")
    return False, None


def money(v):
    return f"{(v or 0):,.2f}"


NAME_HINTS = ("الصنف", "الاسم", "اسم", "البيان", "الطرف", "العميل", "المورد", "الحساب", "الحركة", "البند", "الملاحظات",
              "ملاحظات", "الوصف", "المندوب", "التقرير", "الحقول", "القيمة")
WRAP_AT = 34          # نص أطول من هذا = تفعيل التفاف النص داخل الخلية (صفوف أعلى)
CARD_ROWS_MAX = 300   # أقصى صفوف تُعرض كبطاقات على الموبايل


def _is_num_col(df, c, numeric):
    return c in numeric or pd.api.types.is_numeric_dtype(df[c])


def _cell_text(v, c, money_cols, qty_cols, pct_cols):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if c in pct_cols:
            return f"{v:,.1f}%"
        if c in qty_cols:
            return f"{v:,.3f}".rstrip("0").rstrip(".")
        if c in money_cols:
            return f"{v:,.2f}"
    return str(v)


def _cards_html(df, money_cols, qty_cols, pct_cols, numeric, highlight=frozenset()):
    """نفس بيانات الجدول كبطاقات (تظهر على الموبايل فقط بدلاً من الجدول العريض): لا تمرير أفقي ولا قص للنص."""
    cols = list(df.columns)
    title_col = next((c for c in cols if any(h in str(c) for h in NAME_HINTS)), cols[0])
    out = ['<div class="fz-cards">']
    for _, row in df.head(CARD_ROWS_MAX).iterrows():
        title = html.escape(_cell_text(row[title_col], title_col, money_cols, qty_cols, pct_cols))
        items = []
        for c in cols:
            if c == title_col:
                continue
            v = _cell_text(row[c], c, money_cols, qty_cols, pct_cols)
            if v == "":
                continue
            cls = "n" if _is_num_col(df, c, numeric) else ""
            hl = " hl" if c in highlight else ""
            items.append(f'<div class="fz-kv{hl}"><span class="k">{html.escape(str(c))}</span>'
                         f'<span class="v {cls}">{html.escape(v)}</span></div>')
        out.append(f'<div class="fz-card-row"><div class="t">{title or "—"}</div>{"".join(items)}</div>')
    if len(df) > CARD_ROWS_MAX:
        out.append(f'<div class="fz-cards-more">يُعرض أول {CARD_ROWS_MAX} سطراً من {len(df):,}. '
                   'لعرض الكل افتح الصفحة على شاشة أعرض أو نزّل Excel.</div>')
    out.append("</div>")
    return "".join(out)


def table(df, money_cols=(), qty_cols=(), pct_cols=(), height=None, reverse=True, select_key=None, multi=False,
          cards=None, highlight=()):
    """جدول تفاعلي موحد لكل البرنامج:
    - المحاذاة: العنوان والبيانات بنفس المحاذاة (يمين للعربي) فتقع القيم تحت عناوينها مباشرة.
    - النصوص الطويلة (البند، اسم الحساب، الملاحظات...): أعمدة عريضة + التفاف النص داخل الخلية بدون قص.
    - الترتيب بالنقر على العنوان، وأول عمود منطقي يظهر يميناً.
    - على الموبايل: الجداول للعرض فقط تتحول تلقائياً إلى بطاقات (بدون تمرير أفقي).
    select_key: تفعيل اختيار صف (يرجع أرقام الصفوف المختارة)."""
    if df is None or df.empty:
        st.caption("لا توجد بيانات.")
        return [] if select_key else None
    seen = {}
    cols = []
    for c in df.columns:  # أسماء أعمدة مكررة تكسر التنسيق: نميزها بمسافة غير مرئية
        seen[c] = seen.get(c, 0) + 1
        cols.append(c if seen[c] == 1 else f"{c}{chr(0x200f) * (seen[c] - 1)}")
    df = df.copy()
    df.columns = cols
    numeric = set(money_cols) | set(qty_cols) | set(pct_cols)
    for c in df.columns:   # أعمدة الأرقام رقمية فعلاً (الفراغ يظهر فارغاً لا «None»)، والنصوص بلا «None»
        if c in numeric:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        elif df[c].dtype == object:
            df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else v)
    raw = df
    if reverse:
        df = df[list(reversed(df.columns))]

    def _f(pattern):
        return lambda v: "" if v is None or pd.isna(v) else pattern.format(v)

    fmt = {c: _f("{:,.2f}") for c in money_cols if c in df.columns}
    fmt.update({c: _f("{:,.3f}") for c in qty_cols if c in df.columns})
    fmt.update({c: _f("{:,.1f}%") for c in pct_cols if c in df.columns})
    styler = df.style.format(fmt, na_rep="")
    hl = [c for c in highlight if c in df.columns]
    if hl:   # أعمدة مميزة بالأصفر (مثل Sales Amount و GP)
        styler = styler.set_properties(subset=hl, **{"background-color": "#fef08a", "color": "#111827",
                                                     "font-weight": "bold"})
    config, wrap = {}, False
    text_cols = [c for c in raw.columns if not _is_num_col(raw, c, numeric)]
    main = next((c for c in text_cols if any(h in str(c) for h in NAME_HINTS)), None)
    widths = {}
    for c in df.columns:
        if _is_num_col(df, c, numeric):
            longest = max(len(_cell_text(v, c, set(money_cols), set(qty_cols), set(pct_cols))) for v in df[c].tolist())
            widths[c] = None
        else:
            longest = int(df[c].astype(str).str.len().max() or 0)
            if longest > 30 or (c == main and longest > 16):
                widths[c] = "large"
            elif longest > 20:
                widths[c] = "medium"
            else:
                widths[c] = None
            wrap = wrap or longest > WRAP_AT
        est = {"large": 400, "medium": 200}.get(widths[c]) or (max(longest, len(str(c))) * 8 + 26)
        widths[c] = (widths[c], est)
    # تثبيت عمودي الكود والاسم فقط عندما لا يتسع الجدول للشاشة (حتى يبقيا ظاهرين أثناء التمرير بين الأرقام)
    pin = set()
    if main is not None and sum(e for _, e in widths.values()) > 980:
        pin.add(main)
        first = raw.columns[0]
        if first in text_cols and first != main and int(raw[first].astype(str).str.len().max() or 0) <= 14:
            pin.add(first)
    for c in df.columns:
        config[c] = st.column_config.Column(width=widths[c][0], alignment="right", pinned=True if c in pin else None)
    kwargs = dict(width="stretch", hide_index=True, column_config=config, placeholder="")
    if wrap:
        kwargs["row_height"] = 72   # > 4rem = التفاف النص داخل الخلية (حتى 3 أسطر)
    if height:
        kwargs["height"] = height
    show_cards = (not select_key) if cards is None else cards
    st.session_state["_tbl_seq"] = st.session_state.get("_tbl_seq", 0) + 1
    box = st.container(key=f"fzt_{'c' if show_cards else 'n'}_{st.session_state['_tbl_seq']}")
    with box:
        if select_key:
            event = st.dataframe(styler, key=select_key, on_select="rerun",
                                 selection_mode="multi-row" if multi else "single-row", **kwargs)
            rows = list(event.selection.rows)
        else:
            st.dataframe(styler, **kwargs)
            rows = None
        if show_cards:
            st.markdown(_cards_html(raw, set(money_cols), set(qty_cols), set(pct_cols), numeric, set(highlight)),
                        unsafe_allow_html=True)
    return rows


def open_once(key, value):
    """يفتح نافذة التفاصيل مرة واحدة لكل اختيار جديد (حتى لا تُعاد فتحها مع كل تحديث للصفحة)."""
    if value is None:
        st.session_state.pop(f"_opened_{key}", None)
        return False
    if st.session_state.get(f"_opened_{key}") == value:
        return False
    st.session_state[f"_opened_{key}"] = value
    return True


def build_file(builder, suffix, *args):
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    try:
        builder(path, *args)
        with open(path, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def download_on_demand(label, key, builder, suffix, mime, filename, *args):
    """يجهّز الملف عند الطلب فقط (لا يُبنى PDF مع كل تحديث للصفحة)."""
    store = st.session_state.setdefault("_files", {})
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button(f"⚙️ تجهيز {label}", key=f"prep_{key}"):
            ok, data = run_safe(build_file, builder, suffix, *args)
            if ok:
                store[key] = data
    with c2:
        if key in store:
            st.download_button(f"⬇️ تنزيل {label}", store[key], file_name=filename, mime=mime, key=f"dl_{key}")


def partners(role):
    col = "is_customer" if role == "customer" else "is_supplier"
    with F.db() as conn:
        return conn.execute(f"SELECT partner_code, name FROM business_partners WHERE {col}=1 ORDER BY name").fetchall()


def items():
    with F.db() as conn:
        return conn.execute("SELECT item_code, item_name FROM inventory ORDER BY item_name").fetchall()


def date_range(key, default_days=None, start=None):
    t = date.today()
    d1 = start or (t - timedelta(days=default_days) if default_days else t.replace(month=1, day=1))
    c1, c2 = st.columns(2)
    # القيمة الافتراضية تُمرر فقط إن لم تكن القيمة مستعادة من مسار العمل (منعاً لتحذير تكرار القيمة)
    ka, kb = f"{key}_from", f"{key}_to"
    with c1:
        a = st.date_input("من تاريخ", key=ka, format="YYYY-MM-DD", **({} if ka in st.session_state else {"value": d1}))
    with c2:
        b = st.date_input("إلى تاريخ", key=kb, format="YYYY-MM-DD", **({} if kb in st.session_state else {"value": t}))
    if a > b:
        st.warning("تاريخ البداية بعد تاريخ النهاية.")
    return a.isoformat(), b.isoformat()


# ============================================================
# تسجيل الدخول والجلسة
# ============================================================
def login_page():
    show_flash()
    right, left = st.columns([5, 6], gap="large")
    with right:
        uri, bg = logo_data_uri(420)
        img = f'<img src="{uri}" alt="ELHAGRY"/>' if uri else f'<div class="ar" style="font-size:2.4rem">{F.BRAND_AR}</div>'
        st.markdown(f'<div class="fz-login-panel" style="background:radial-gradient(circle at 50% 30%, #13306b 0%, {bg} 62%)">'
                    f'{img}<div class="ar">{F.BRAND_AR}</div><div class="en">{F.BRAND_EN.upper()}</div></div>',
                    unsafe_allow_html=True)
    with left:
        st.markdown(f'<div class="fz-login-title">تسجيل الدخول</div>'
                    f'<div class="fz-login-sub">أهلاً بك في {F.BRAND_AR}</div>', unsafe_allow_html=True)
        with st.form("login", clear_on_submit=False):
            username = st.text_input("اسم المستخدم", autocomplete="username")
            password = st.text_input("كلمة المرور", type="password", autocomplete="current-password")
            ok = st.form_submit_button("دخول", width="stretch", type="primary")
        st.caption("لأمان النظام لا تُستعاد كلمة المرور تلقائياً؛ تواصل مع مدير النظام لإعادة تعيينها.")
        if ok:
            try:
                user = F.authenticate(username, password)
            except F.UserError as e:
                st.error(str(e))
                return
            st.session_state.user = user
            st.session_state.last_seen = time.time()
            st.rerun()


def check_session():
    if not U():
        return False
    now = time.time()
    if now - st.session_state.get("last_seen", now) > SESSION_TIMEOUT_MIN * 60:
        st.session_state.clear()
        flash(f"انتهت الجلسة بعد {SESSION_TIMEOUT_MIN} دقيقة من عدم النشاط. سجّل الدخول مرة أخرى.", "info")
        return False
    st.session_state.last_seen = now
    return True


def header(title):
    u = U()
    st.markdown(f'<div class="fz-header"><div><div class="t">{html.escape(title)}</div>'
                f'<div class="b">{F.BRAND_FULL}</div></div>'
                f'<div class="u">👤 {html.escape(u["full_name"])} • {html.escape(u["username"])}</div></div>',
                unsafe_allow_html=True)
    show_flash()


# ============================================================
# لوحة المؤشرات
# ============================================================
def page_dashboard():
    header("🏠 لوحة المؤشرات")
    quick = [(k, t) for k, t in (("import", "📥 مركز استيراد البيانات"), ("purchase.discount", "🏷️ خصم لاحق من مورد"),
                                  ("sales.discount", "🏷️ إشعار دائن لعميل"), ("discounts", "📒 سجل الخصومات اللاحقة"))
             if can(k)]
    if quick:
        cols = st.columns(len(quick))
        for c, (k, t) in zip(cols, quick):
            if k.endswith(".discount"):
                c.button(t, key=f"quick_{k}", width="stretch", on_click=go_discount, args=(k.split(".")[0],))
            else:
                c.button(t, key=f"quick_{k}", width="stretch", on_click=go_to, args=(k,))
    ok, m = run_safe(F.fetch_dashboard_metrics)
    if not ok:
        return
    fin = can_fin("customer") and can_fin("supplier")
    rows = [
        [("صافي المبيعات (قبل الضريبة)", m["sales"]), ("مجمل الربح", m["gross"]), ("صافي الربح", m["profit"]),
         ("المصروفات التشغيلية", m["expenses"])],
        [("صافي المشتريات", m["purchases"]), ("قيمة المخزون", m["stock"]), ("الصناديق", m["cash_box"]),
         ("البنوك", m["banks"])],
    ]
    if fin:
        rows.append([("الذمم المدينة (العملاء)", m["recv"]), ("الذمم الدائنة (الموردين)", m["payb"]),
                     ("صافي المركز", m["net_position"]), ("مصروف الإهلاك", m["depreciation"])])
    for r in rows:
        cols = st.columns(len(r))
        for c, (label, v) in zip(cols, r):
            c.metric(label, f"{v:,.2f} ر.س")
    best = [(t, m[k]) for t, k in (("أفضل عميل", "best_cust"), ("أفضل مورد", "best_supp"), ("الصنف الأكثر مبيعاً", "best_item"))
            if m.get(k)]
    if best:
        st.caption("   |   ".join(f"{t}: {v[0]} ({v[1]:,.2f})" for t, v in best))

    ok, ch = run_safe(F.fetch_dashboard_charts, 6)
    if ok:
        keys, sales, purch = ch["flow"]
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### المبيعات مقابل المشتريات (آخر 6 أشهر)")
            df = pd.DataFrame({"الشهر": [F.AR_MONTHS[int(k[5:7]) - 1] + " " + k[:4] for k in keys],
                               "المبيعات": sales, "المشتريات": purch}).set_index("الشهر")
            st.bar_chart(df, color=[BLUE, TEAL], stack=False)
        with c2:
            st.markdown("#### المصروفات التشغيلية (السنة الحالية)")
            if ch["exp"]:
                df = pd.DataFrame([(n or "بنود أخرى", v) for n, v in ch["exp"]], columns=["البند", "المبلغ"]).set_index("البند")
                st.bar_chart(df, color=CYAN, horizontal=True)
            else:
                st.caption("لا توجد مصروفات مسجلة هذه السنة.")
        if ch["cash"]:
            st.markdown("#### أرصدة الخزينة والبنوك")
            table(pd.DataFrame(ch["cash"], columns=["الحساب", "الرصيد"]), money_cols=["الرصيد"])
    if can("expiry"):
        ok, ex = run_safe(F.fetch_expiry_summary)
        if ok and (ex["expired"] or ex["near"]):
            st.warning(f"⏳ تنبيه الصلاحية: {len(ex['expired'])} دفعة منتهية (قيمة {money(ex['expired_value'])}) "
                       f"و {len(ex['near'])} دفعة تنتهي خلال {ex['alert_days']} يوماً (قيمة {money(ex['near_value'])}).")


# ============================================================
# المبيعات والمشتريات + الخصم اللاحق
# ============================================================
def invoices_table(kind):
    c = F.FLOW[kind]
    sales = kind == "sales"
    d1, d2 = date_range(f"inv_{kind}", default_days=90)
    with F.db() as conn:
        rows = conn.execute(
            f"SELECT h.{c['iid']}, COALESCE(h.invoice_no, h.{c['iid']}), h.{c['idate']}, h.{c['pcode']}, h.{c['pname']}, "
            f"h.subtotal, h.vat_amount, h.total_amount, h.paid_amount, "
            f"COALESCE((SELECT SUM(total_amount) FROM post_discounts d WHERE d.kind=? AND d.invoice_id=h.{c['iid']}),0), "
            f"h.status, COALESCE(h.source_type,'مباشر') FROM {c['inv']} h WHERE h.{c['idate']} BETWEEN ? AND ? "
            f"ORDER BY h.{c['idate']} DESC, h.{c['iid']} DESC", (kind, d1, d2)).fetchall()
    att = F.attachment_counts("purchase") if not sales else {}
    data = [(r[1], r[2], f"{r[3]} - {r[4]}", r[5], r[6], r[7], r[8], r[9], max(r[7] - r[8] - r[9], 0), r[10], r[11])
            + (() if sales else (f"📎 {att[r[0]]}" if att.get(r[0]) else "-",)) for r in rows]
    cols = ["رقم الفاتورة", "التاريخ", "العميل" if sales else "المورد", "قبل الضريبة", "الضريبة", "الإجمالي", "المسدد",
            "خصومات لاحقة", "المتبقي", "الحالة", "المصدر"] + ([] if sales else ["المرفقات"])
    df = pd.DataFrame(data, columns=cols)
    if not can_fin("customer" if sales else "supplier") and not df.empty:
        df = df.drop(columns=["المسدد", "خصومات لاحقة", "المتبقي"])
    st.caption("🖱️ اختر أي فاتورة من الجدول لعرض تفاصيلها وبنودها أسفل الجدول مباشرة (والنقر على عنوان العمود يرتّب الجدول).")
    sel = table(df, money_cols=["قبل الضريبة", "الضريبة", "الإجمالي", "المسدد", "خصومات لاحقة", "المتبقي"],
                select_key=f"inv_tbl_{kind}")
    if rows:
        st.caption(f"عدد الفواتير: {len(rows)}   |   الإجمالي: {money(sum(r[7] for r in rows))} ر.س")
    picked = rows[sel[0]][0] if sel else None
    if open_once(f"inv_{kind}", picked):
        st.session_state[f"inv_inline_{kind}"] = picked
    inline_invoice(kind)


def inline_invoice(kind):
    """تفاصيل الفاتورة المختارة داخل الصفحة نفسها (بدون نوافذ منبثقة): تبقى ظاهرة حتى يغلقها المستخدم."""
    inv_id = st.session_state.get(f"inv_inline_{kind}")
    if not inv_id:
        return
    with st.container(border=True, key=f"inv_inline_box_{kind}"):
        st.button("✖ إغلاق تفاصيل الفاتورة", key=f"inv_inline_x_{kind}", type="tertiary",
                  on_click=set_state, kwargs={f"inv_inline_{kind}": None})
        invoice_body(kind, inv_id)


def invoice_body(kind, inv_id, on_pay=None):
    """تفاصيل الفاتورة (البنود والضريبة والمتبقي) + السداد. on_pay: ماذا يفعل زر «سداد هذه الفاتورة»
    (داخل مساحة العمل يفتح تبويب السداد المجاور، وخارجها ينقل لشاشة الخزينة مع زر رجوع للفاتورة)."""
    c = F.FLOW[kind]
    sales = kind == "sales"
    with F.db() as conn:
        h = conn.execute(f"SELECT COALESCE(invoice_no, {c['iid']}), {c['idate']}, {c['pcode']}, {c['pname']}, subtotal, "
                         f"vat_amount, total_amount, paid_amount, status, COALESCE(notes,'') FROM {c['inv']} WHERE {c['iid']}=?",
                         (inv_id,)).fetchone()
    if not h:
        st.error("الفاتورة غير موجودة.")
        return
    st.markdown(f"### {'فاتورة ضريبية' if sales else 'فاتورة مشتريات'} {h[0]}")
    st.caption(f"{h[1]}   |   {h[2]} - {h[3]}   |   الحالة: {h[8]}" + (f"   |   {h[9]}" if h[9] else ""))
    m = st.columns(4)
    m[0].metric("قبل الضريبة", money(h[4]))
    m[1].metric("الضريبة 15%", money(h[5]))
    m[2].metric("الإجمالي", money(h[6]))
    m[3].metric("المسدد", money(h[7]))
    show_invoice_lines(kind, inv_id)
    pre = F.invoice_payment_prefill(kind, inv_id)
    if pre["remaining"] > 0.005:
        st.info(f"المتبقي على الفاتورة بعد المسدد والخصومات اللاحقة: {money(pre['remaining'])} ر.س")
        if can("cash") and can("cash.receipt" if sales else "cash.payment"):
            st.button(f"💳 سداد هذه الفاتورة ({money(pre['remaining'])} ر.س)", type="primary", width="stretch",
                      key=f"dlg_pay_{kind}_{inv_id}", on_click=on_pay or go_pay_invoice, args=(kind, inv_id))
    else:
        st.success("✅ الفاتورة مسددة بالكامل.")
    b1, b2 = st.columns(2)
    if can(f"{kind}.discount"):
        b1.button("🏷️ تسجيل خصم لاحق / إشعار دائن", type="primary", key=f"dlg_disc_{kind}_{inv_id}",
                  on_click=go_discount, args=(kind, inv_id))
    with b2:
        download_on_demand("PDF", f"inv_{kind}_{inv_id}", F.build_invoice_pdf, ".pdf", "application/pdf",
                           f"invoice_{h[0]}.pdf", kind, inv_id)
    if not sales:
        attachments_box("purchase", inv_id)


def attachments_box(doc_type, doc_id):
    """مرفقات المستند: عرض الصور مباشرة، تنزيل الملفات، ورفع ملفات جديدة."""
    st.markdown("#### 📎 المرفقات (فاتورة المورد الأصلية / المستندات المؤيدة)")
    for aid, name, size, mime, by, at in F.list_attachments(doc_type, doc_id):
        _n, _m, data = F.get_attachment(aid)
        c1, c2 = st.columns([3, 1])
        c1.caption(f"{name} • {size / 1024:,.1f} KB • {by or '-'} • {at}")
        c2.download_button("⬇️ فتح / تنزيل", data, file_name=name, mime=mime, key=f"att_dl_{aid}")
        if (mime or "").startswith("image/"):
            st.image(data, width=420)
    up = st.file_uploader("رفع صورة أو PDF", type=[e.lstrip(".") for e in F.ATTACH_EXTS], key=f"att_up_{doc_type}_{doc_id}",
                          accept_multiple_files=True)
    if up and st.button("📤 حفظ المرفقات", key=f"att_save_{doc_type}_{doc_id}", type="primary"):
        for f_ in up:
            run_safe(F.add_attachment, doc_type, doc_id, f_.name, f_.getvalue(), U()["username"])
        st.session_state.pop(f"att_up_{doc_type}_{doc_id}", None)
        st.rerun()


def show_invoice_lines(kind, inv_id):
    with F.db() as conn:
        if kind == "sales":
            rows = conn.execute("SELECT item_code, item_name, quantity, unit, unit_price, COALESCE(discount_amount,0), "
                                "subtotal, vat_amount, total_amount FROM sales_invoice_lines WHERE invoice_id=? "
                                "ORDER BY line_id", (inv_id,)).fetchall()
            cols = ["الكود", "الصنف", "الكمية", "الوحدة", "السعر", "الخصم الخاص", "الخاضع للضريبة", "الضريبة", "الإجمالي"]
        else:
            rows = conn.execute("SELECT item_code, item_name, quantity, unit, unit_cost, subtotal, vat_amount, total_amount, "
                                "COALESCE(production_date,''), COALESCE(expiry_date,'') FROM purchase_invoice_lines "
                                "WHERE purchase_id=? ORDER BY line_id", (inv_id,)).fetchall()
            cols = ["الكود", "الصنف", "الكمية", "الوحدة", "التكلفة", "قبل الضريبة", "الضريبة", "الإجمالي", "الإنتاج", "الانتهاء"]
    table(pd.DataFrame(rows, columns=cols), money_cols=["السعر", "الخصم الخاص", "الخاضع للضريبة", "الضريبة", "الإجمالي",
                                                         "التكلفة", "قبل الضريبة"], qty_cols=["الكمية"])


def unbilled_section(kind):
    """كل الأذون غير المفوترة (مبيعات + مشتريات) في جدول واحد، الأقدم أولاً، مع بحث وتصفية وإصدار الفاتورة."""
    c1, c2 = st.columns([2, 1])
    q = c1.text_input("🔎 بحث (الطرف / رقم الإذن / رقم الأمر)", key=f"ub_q_{kind}")
    typ = c2.selectbox("النوع", ["الكل", "المبيعات", "المشتريات"], key=f"ub_t_{kind}")
    rows = F.all_unbilled({"المبيعات": "sales", "المشتريات": "purchase"}.get(typ), q)
    df = pd.DataFrame([("مبيعات (تسليم)" if r["kind"] == "sales" else "مشتريات (استلام)", r["id"], r["date"], r["age"],
                        r["order"], f"{r['party_code']} - {r['party']}", r["lines"], r["subtotal"], r["vat"], r["total"])
                       for r in rows],
                      columns=["النوع", "رقم الإذن", "التاريخ", "منذ (يوم)", "رقم الأمر", "الطرف", "البنود", "قبل الضريبة",
                               "الضريبة", "الإجمالي"])
    st.caption(f"عدد الأذون غير المفوترة: {len(rows)}   |   الإجمالي شامل الضريبة: {money(sum(r['total'] for r in rows))} ر.س"
               "   |   الأقدم أولاً حتى لا يُنسى شيء.")
    sel = table(df, money_cols=["قبل الضريبة", "الضريبة", "الإجمالي"], select_key=f"ub_tbl_{kind}", multi=True)
    picked = [rows[i] for i in sel]
    if picked and (can("sales.invoices") or can("purchase.invoices")):
        kinds = {r["kind"] for r in picked}
        parties = {r["party_code"] for r in picked}
        if len(kinds) > 1 or len(parties) > 1:
            st.warning("الفاتورة الواحدة تكون من نوع واحد ولطرف واحد. عدّل التحديد.")
            return
        k = kinds.pop()
        if not can(f"{k}.invoices"):
            st.warning("ليس لديك صلاحية إصدار هذا النوع من الفواتير.")
            return
        c1, c2, c3 = st.columns(3)
        d = c1.date_input("تاريخ الفاتورة", date.today(), key=f"ub_d_{kind}", format="YYYY-MM-DD")
        notes = c2.text_input("ملاحظات", key=f"ub_n_{kind}")
        ref = c3.text_input("رقم فاتورة المورد", key=f"ub_r_{kind}") if k == "purchase" else ""
        total = sum(r["total"] for r in picked)
        if k == "sales":
            cst = F.credit_status(picked[0]["party_code"], 0.0, "invoice")
            if cst["exceeded"]:
                st.warning("⚠ " + cst["message"].replace("\n", "  \n") + "  \n(الفوترة من أذون التسليم لا تُمنع؛ "
                           "البضاعة خرجت والفاتورة الضريبية إلزامية.)")
        if st.button(f"🧾 إصدار فاتورة {'ضريبية' if k == 'sales' else 'مورد'} من {len(picked)} إذن ({money(total)} ر.س)",
                     type="primary", key=f"ub_go_{kind}"):
            ok, res = run_safe(F.invoice_unbilled_selection, k, [r["id"] for r in picked], d.isoformat(), notes, ref,
                               allow_closed=U()["is_admin"])
            if ok:
                flash(f"تم إصدار الفاتورة رقم {F.invoice_number(k, res[0])} بإجمالي {money(res[1])} ر.س")
                st.rerun()


def discount_section(kind):
    """الخصم اللاحق: إشعار دائن للعميل / إشعار خصم من المورد (نفس منطق واجهة سطح المكتب)."""
    sales = kind == F.DISC_SALES
    role = "customer" if sales else "supplier"
    who = "العميل" if sales else "المورد"
    st.info("إشعار دائن للعميل: خصم أو تسوية بعد إصدار الفاتورة وتسليم البضاعة. يخفض ما على العميل وصافي المبيعات "
            "وضريبة المخرجات، ولا يغير المخزون." if sales else
            "إشعار خصم من المورد بعد دخول البضاعة المخزن: يخفض ما للمورد علينا وضريبة المدخلات. «تخفيض تكلفة المخزون» "
            "يوزع الخصم على أصناف الفاتورة ويخفض متوسط تكلفتها (وما يخص بضاعة سبق بيعها يظهر خصماً مكتسباً)، "
            "و«خصم مباشر» يُسجَّل خصماً مكتسباً بقائمة الدخل.")
    records = F.list_post_discounts(kind)
    can_mod = can("modify")
    labels = {r[0]: f"{r[1]} | {r[3]} | {r[5]} | {money(r[10])}" for r in records}
    edit_id = None
    if can_mod and records:
        edit_id = st.selectbox("الإشعار", [None] + [r[0] for r in records],
                               format_func=lambda v: "➕ إشعار جديد" if v is None else f"✏️ تعديل {labels[v]}",
                               key=f"disc_mode_{kind}")
    rec, rec_lines = (F.load_post_discount(edit_id) if edit_id else (None, []))
    sfx = f"{kind}_{edit_id or 'new'}"

    plist = partners(role)
    if not plist:
        st.warning(f"لا يوجد {who} مسجل.")
        return
    names = dict(plist)
    codes = [c for c, _ in plist]
    idx = codes.index(rec["party_code"]) if rec and rec["party_code"] in codes else None
    c1, c2 = st.columns(2)
    with c1:
        party = st.selectbox(who, codes, index=idx, format_func=lambda c: f"{c} - {names[c]}",
                             placeholder=f"اختر {who}...", key=f"disc_party_{sfx}")
    invs = F.discount_invoice_choices(kind, party) if party else []
    inv_map = {r[0]: r for r in invs}
    with c2:
        inv_opts = [None] + [r[0] for r in invs]
        inv_idx = inv_opts.index(rec["invoice_id"]) if rec and rec["invoice_id"] in inv_opts else 0
        invoice = st.selectbox("الفاتورة المرتبطة (اختياري)", inv_opts, index=inv_idx, key=f"disc_inv_{sfx}",
                               format_func=lambda v: "بدون فاتورة محددة (على الحساب)" if v is None else
                               f"{inv_map[v][1]} ({inv_map[v][2]}) - قبل الضريبة {money(inv_map[v][3])} - "
                               f"خصومات سابقة {money(inv_map[v][5])}")
    c1, c2, c3 = st.columns(3)
    with c1:
        d = st.date_input("التاريخ", date.fromisoformat(rec["discount_date"]) if rec else date.today(),
                          key=f"disc_date_{sfx}", format="YYYY-MM-DD")
    with c2:
        amount = st.number_input("قيمة الخصم (قبل الضريبة)", min_value=0.0, step=10.0, format="%.2f",
                                 value=float(rec["subtotal"]) if rec else 0.0, key=f"disc_amt_{sfx}")
    with c3:
        apply_vat = st.checkbox("تسوية ضريبة القيمة المضافة 15%", value=(rec["vat_amount"] > 0.005) if rec else True,
                                key=f"disc_vat_{sfx}")
    method, item = F.DISC_PL, None
    if not sales:
        c1, c2 = st.columns(2)
        with c1:
            m_opts = [F.DISC_STOCK, F.DISC_PL]
            method = st.radio("طريقة المعالجة", m_opts, index=m_opts.index(rec["method"]) if rec else 0,
                              format_func=lambda v: F.DISC_METHOD_LABEL[v], key=f"disc_method_{sfx}", horizontal=True)
        with c2:
            if method == F.DISC_STOCK:
                if invoice:
                    with F.db() as conn:
                        ilist = conn.execute("SELECT DISTINCT item_code, item_name FROM purchase_invoice_lines "
                                             "WHERE purchase_id=?", (invoice,)).fetchall()
                else:
                    ilist = items()
                inames = dict(ilist)
                i_opts = [None] + [c for c, _ in ilist]
                pre = rec_lines[0][0] if rec and rec["method"] == F.DISC_STOCK and len(rec_lines) == 1 else None
                item_code = st.selectbox("على صنف محدد" + (" (اختياري: الافتراضي توزيع على كل أصناف الفاتورة)" if invoice
                                                            else " (مطلوب بدون فاتورة)"),
                                         i_opts, index=i_opts.index(pre) if pre in i_opts else 0, key=f"disc_item_{sfx}",
                                         format_func=lambda v: "كل أصناف الفاتورة" if v is None else f"{v} - {inames[v]}")
                item = (item_code, inames[item_code]) if item_code else None
    c1, c2 = st.columns(2)
    with c1:
        reason = st.text_input("سبب الخصم", value=rec["reason"] if rec else "", key=f"disc_reason_{sfx}")
    with c2:
        notes = st.text_input("ملاحظات", value=rec["notes"] if rec else "", key=f"disc_notes_{sfx}")

    vat = round(amount * F.VAT_RATE, 2) if apply_vat else 0.0
    if invoice and (inv_map[invoice][4] - inv_map[invoice][3]) <= 0.005:
        vat = 0.0
    total = round(amount + vat, 2)
    if party and can_fin(role):
        ok, stmt = run_safe(F.fetch_partner_statement, party)
        if ok:
            bal = stmt[4]
            if rec and rec["party_code"] == party:
                bal += rec["total_amount"] if kind == F.DISC_PURCHASE else -rec["total_amount"]
            new_bal = bal + total if kind == F.DISC_PURCHASE else bal - total
            side = lambda b: "مدين (عليه)" if b > 0.005 else ("دائن (له)" if b < -0.005 else "")
            st.caption(f"الضريبة: {money(vat)}   |   إجمالي الإشعار: {money(total)} ر.س   |   "
                       f"رصيد {who} الحالي: {money(abs(bal))} {side(bal)}   ←   بعد الإشعار: {money(abs(new_bal))} {side(new_bal)}")
    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("💾 حفظ الإشعار وترحيله", type="primary", key=f"disc_save_{sfx}", width="stretch"):
            if not party:
                st.error(f"اختر {who}.")
            else:
                ok, res = run_safe(F.save_post_discount, kind, d.isoformat(), (party, names[party]), amount,
                                   apply_vat=apply_vat, invoice_id=invoice, method=method, item=item, reason=reason.strip(),
                                   notes=notes.strip(), discount_id=edit_id, user=U()["username"],
                                   allow_closed=U()["is_admin"])
                if ok:
                    flash(f"تم {'تعديل' if edit_id else 'تسجيل'} الإشعار رقم {res[1]} بإجمالي {money(res[2])} ر.س "
                          f"وترحيل القيد اليومي الآلي.")
                    st.session_state.pop(f"disc_mode_{kind}", None)
                    st.rerun()
    with b2:
        if edit_id:
            confirm = st.checkbox("تأكيد حذف الإشعار", key=f"disc_delc_{sfx}")
            if st.button("🗑️ حذف الإشعار", disabled=not confirm, key=f"disc_del_{sfx}", width="stretch"):
                ok, _ = run_safe(F.delete_post_discount, edit_id, allow_closed=U()["is_admin"])
                if ok:
                    flash("تم حذف الإشعار وإعادة الأرصدة والتكلفة وحذف قيده الآلي.")
                    st.session_state.pop(f"disc_mode_{kind}", None)
                    st.rerun()
    if rec and rec_lines:
        st.markdown("**توزيع الخصم على الأصناف**")
        table(pd.DataFrame([(c, n, a, ap, a - ap) for c, n, a, ap in rec_lines],
                           columns=["الكود", "الصنف", "الخصم الموزع", "خُفّض من تكلفة المخزون", "خصم مكتسب (بضاعة مباعة)"]),
              money_cols=["الخصم الموزع", "خُفّض من تكلفة المخزون", "خصم مكتسب (بضاعة مباعة)"])

    st.markdown("#### سجل الإشعارات")
    df = pd.DataFrame([(r[1], r[3], f"{r[4]} - {r[5]}", r[6] or "-",
                        ("تخفيض التكلفة" if r[7] == F.DISC_STOCK else "خصم مباشر") if not sales else "إشعار دائن",
                        r[8], r[9], r[10], r[11] if r[7] == F.DISC_STOCK else 0.0, r[12] or "-") for r in records],
                      columns=["رقم الإشعار", "التاريخ", who, "الفاتورة", "المعالجة", "قبل الضريبة", "الضريبة", "الإجمالي",
                               "خُفّض من المخزون", "السبب"])
    if sales and not df.empty:
        df = df.drop(columns=["خُفّض من المخزون"])
    table(df, money_cols=["قبل الضريبة", "الضريبة", "الإجمالي", "خُفّض من المخزون"])
    if records:
        st.caption(f"عدد الإشعارات: {len(records)}   |   الإجمالي شامل الضريبة: {money(sum(r[10] for r in records))} ر.س")
        pr = st.selectbox("🖨️ طباعة إشعار", [r[0] for r in records], format_func=lambda v: labels[v],
                          key=f"disc_print_{kind}")
        if pr:
            download_on_demand("الإشعار الدائن PDF" if sales else "إشعار خصم المورد PDF", f"dn_{pr}",
                               F.build_post_discount_pdf, ".pdf", "application/pdf", f"credit_note_{pr}.pdf", pr)


def go_to(page, **state):
    """الانتقال لصفحة أخرى مع ضبط حالة اختيارية. تُستدعى كـ on_click (قبل رسم الصفحة) حتى يُسمح بضبط حالة العناصر."""
    st.session_state["nav"] = page
    for k, v in state.items():
        if v is None:
            st.session_state.pop(k, None)
        else:
            st.session_state[k] = v


def set_state(**state):
    """ضبط حالة بدون تغيير الصفحة (callback)."""
    for k, v in state.items():
        if v is None:
            st.session_state.pop(k, None)
        else:
            st.session_state[k] = v


# ============================================================
# 🗂️ مساحة العمل: تبويبات متجاورة حقيقية (st.tabs) داخل الشاشة الرئيسية — بدون أي نوافذ منبثقة
# التقرير | كشف حساب الطرف | تفاصيل الفاتورة | سند السداد: كلها ظاهرة في نفس الشاشة ويُرسم محتواها في كل تحديث،
# فلا تُحذف حالة أي عنصر (الفلاتر، الطرف، الفاتورة، بيانات السند) عند التنقل بين التبويبات.
# ============================================================
WS_KINDS = ("report", "statement", "invoice", "pay")
WS_HOME = {"reports": "report", "statement": "statement"}

# مفاتيح تبقى محفوظة حتى لو لم تُرسم عناصرها في هذا التحديث (Streamlit يحذف حالة العنصر غير المرسوم افتراضياً)
PERSIST_KEYS = ("rep_group", "rep_key", "rep_from", "rep_to", "rep_grp", "smr_rep", "smr_cust",
                "stmt_party", "stmt_scope", "stmt_from", "stmt_to", "stmt_q", "stmt_hide0",
                "flow_view_sales", "flow_view_purchase",
                "cv_type", "cv_bank", "cv_cat", "cv_party", "cv_inv", "cv_amount", "cv_desc")


def keep_state_alive():
    for k in PERSIST_KEYS:
        if k in st.session_state:
            st.session_state[k] = st.session_state[k]


def _party_name(code):
    with F.db() as conn:
        r = conn.execute("SELECT name FROM business_partners WHERE partner_code=?", (code,)).fetchone()
    return r[0] if r else str(code)


def _ws_available(kind):
    if kind == "report":
        return can("reports")
    if kind == "statement":
        return can("statement")
    if kind == "invoice":
        v = st.session_state.get("inv_view")
        return bool(v) and (can(v[0]) or can("statement") or can("reports"))
    if kind == "pay":
        v = st.session_state.get("cash_pay")
        return bool(v) and can("cash") and can("cash.receipt" if v[0] == "sales" else "cash.payment")
    return False


def _ws_label(kind):
    try:
        if kind == "report":
            k = st.session_state.get("rep_key")
            return "📊 " + (F.REPORT_INDEX[k][0] if k in F.REPORT_INDEX else "التقرير")
        if kind == "statement":
            p = st.session_state.get("stmt_party")
            return "📜 كشف " + (_party_name(p) if p else "الحساب")
        if kind == "invoice":
            k, i = st.session_state["inv_view"]
            return f"🧾 فاتورة {'مبيعات' if k == 'sales' else 'مشتريات'} {F.invoice_number(k, i)}"
        if kind == "pay":
            k, i = st.session_state["cash_pay"]
            return f"💳 سداد {F.invoice_number(k, i)}"
    except Exception:  # noqa: BLE001  (عنوان التبويب لا يكسر الصفحة)
        pass
    return {"report": "📊 التقرير", "statement": "📜 كشف الحساب", "invoice": "🧾 الفاتورة", "pay": "💳 السداد"}[kind]


def ws_open_tab(kind, **state):
    """فتح (أو تحديث) تبويب مجاور في مساحة العمل وتفعيله، مع بقاء كل التبويبات الأخرى كما هي.
    يعمل كـ callback أو داخل الرسم (يتبعه st.rerun)."""
    set_state(**state)
    opened = list(st.session_state.get("ws_open") or [])
    home = WS_HOME.get(st.session_state.get("nav"))
    for k in (home, kind):
        if k and k not in opened:
            opened.append(k)
    st.session_state["ws_open"] = opened
    st.session_state["ws_active"] = kind


def ws_close(kind):
    st.session_state["ws_open"] = [k for k in (st.session_state.get("ws_open") or []) if k != kind]
    st.session_state["ws_active"] = WS_HOME.get(st.session_state.get("nav"))
    if kind == "pay":
        for k in ("cash_pay", "cash_pay_info"):
            st.session_state.pop(k, None)


def ws_close_all():
    for k in ("ws_open", "ws_active", "cash_pay", "cash_pay_info", "inv_view"):
        st.session_state.pop(k, None)


def _ws_on_tab():
    st.session_state["ws_active"] = (st.session_state.get("_ws_labels") or {}).get(st.session_state.get("ws_tab"))


def ws_pay(kind, inv_id):
    """«سداد هذه الفاتورة» داخل مساحة العمل: تجهيز السند كاملاً في الحالة ثم فتح تبويب السداد المجاور."""
    st.session_state["cash_prefill"] = F.invoice_payment_prefill(kind, inv_id)
    ws_open_tab("pay", cash_pay=(kind, inv_id))


def workspace(home):
    """يرسم الصفحة كتبويبات متجاورة: تبويب الصفحة الأساسي + ما فتحه المستخدم بالتعمق (كشف/فاتورة/سداد)."""
    opened = set(st.session_state.get("ws_open") or []) | {home}
    kinds = [k for k in WS_KINDS if k in opened and (k == home or _ws_available(k))]
    if len(kinds) == 1:
        _ws_body(home)
        return
    labels = {k: _ws_label(k) for k in kinds}
    active = st.session_state.get("ws_active")
    if active not in kinds:
        active = home
    st.session_state["ws_active"] = active
    st.session_state["_ws_labels"] = {v: k for k, v in labels.items()}
    st.session_state["ws_tab"] = labels[active]
    with st.container(horizontal=True, key="ws_bar", vertical_alignment="center"):
        st.caption("🗂️ مساحة العمل: كل خطوة في تبويب مستقل — تنقّل بينها بحرية دون فقد أي بيانات.")
        st.button("🧹 إغلاق التبويبات", key="ws_close_all", type="tertiary", on_click=ws_close_all)
    tabs = st.tabs([labels[k] for k in kinds], key="ws_tab", on_change=_ws_on_tab)
    for k, tab in zip(kinds, tabs):
        with tab:
            if k != home:
                st.button("✖ إغلاق هذا التبويب", key=f"ws_x_{k}", type="tertiary", on_click=ws_close, args=(k,))
            _ws_body(k)


def _ws_body(kind):
    if kind == "report":
        reports_body()
    elif kind == "statement":
        statement_body()
    elif kind == "invoice":
        invoice_body(*st.session_state["inv_view"], on_pay=ws_pay)
    elif kind == "pay":
        voucher_form(in_workspace=True)


def go_discount(kind, inv_id=None):
    """فتح «تسجيل خصم لاحق / إشعار دائن» داخل المبيعات/المشتريات مع اختيار الطرف والفاتورة تلقائياً."""
    party = None
    if inv_id:
        c = F.FLOW[kind]
        with F.db() as conn:
            r = conn.execute(f"SELECT {c['pcode']} FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()
        party = r[0] if r else None
    for k in (f"disc_mode_{kind}", f"disc_party_{kind}_new", f"disc_inv_{kind}_new"):
        st.session_state.pop(k, None)
    go_to(kind, **{f"flow_view_{kind}": "disc", f"disc_party_{kind}_new": party, f"disc_inv_{kind}_new": inv_id})


def page_flow(kind):
    sales = kind == "sales"
    header("🛒 المبيعات" if sales else "📦 المشتريات")
    views = []
    if can(f"{kind}.orders"):
        views.append(("ord", "📝 1) الأوامر"))
    if can(f"{kind}.{'delivery' if sales else 'receipt'}"):
        views.append(("ful", "🚚 2) أذون التسليم" if sales else "📥 2) أذون الاستلام"))
    if can(f"{kind}.invoices"):
        views.append(("review", "✅ 3) مراجعة واعتماد الفاتورة"))
    if can(f"{kind}.invoices") or can(f"{kind}.direct"):
        views.append(("inv", "🧾 الفواتير"))
    if can(f"{kind}.invoices"):
        views.append(("unbilled", "📋 غير المفوتر (مبيعات + مشتريات)"))
    if can(f"{kind}.discount"):
        views.append(("disc", "🏷️ تسجيل خصم لاحق / إشعار دائن"))
    if not views:
        st.warning("لا توجد صلاحيات تفصيلية متاحة لك في هذه الصفحة.")
        return
    labels = dict(views)
    key = f"flow_view_{kind}"
    pending = st.session_state.pop(f"_flow_goto_{kind}", None)
    if pending in labels:
        st.session_state[key] = pending
    if st.session_state.get(key) not in labels:
        st.session_state.pop(key, None)
    view = st.radio("القسم", list(labels), format_func=labels.get, horizontal=True, key=key,
                    label_visibility="collapsed")
    if view == "ord":
        orders_section(kind)
    elif view == "ful":
        fulfil_section(kind)
    elif view == "review":
        invoice_review_section(kind)
    elif view == "inv":
        st.caption("الدورة المستندية إجبارية: أمر ← تسليم/استلام ← فاتورة ضريبية (لا توجد فواتير مباشرة بدون أمر).")
        opened = st.session_state.pop(f"inv_open_{kind}", None)
        if opened:
            st.session_state[f"inv_inline_{kind}"] = opened
        invoices_table(kind)
    elif view == "unbilled":
        unbilled_section(kind)
    else:
        discount_section(kind)


# ============================================================
# المخزون وكارت الصنف
# ============================================================
def page_inventory():
    header("🏭 المخزون وكارت الصنف")
    t1, t2 = st.tabs(["📦 أرصدة الأصناف", "🗂️ كارت الصنف"])
    with t1:
        q = st.text_input("بحث بالكود أو الاسم أو الباركود أو الكود القديم", key="inv_q")
        with F.db() as conn:
            rows = conn.execute("SELECT item_code, item_name, unit, quantity, cost_price, quantity*cost_price, "
                                "COALESCE(selling_price,0), COALESCE(barcode,''), COALESCE(legacy_code,'') "
                                "FROM inventory ORDER BY item_name").fetchall()
        if q:
            nq = F.search_norm(q)
            rows = [r for r in rows if nq in F.search_norm(f"{r[0]} {r[1]} {r[7]} {r[8]}")]
        df = pd.DataFrame([r[:7] for r in rows], columns=["الكود", "الصنف", "الوحدة", "الرصيد", "متوسط التكلفة", "القيمة",
                                                          "سعر البيع"])
        st.caption("🖱️ اختر أي صنف لعرض تفاصيله وآخر حركاته في نافذة منبثقة. الأكواد الجديدة تُولَّد تلقائياً (ELH-XXXX).")
        sel = table(df, money_cols=["متوسط التكلفة", "القيمة", "سعر البيع"], qty_cols=["الرصيد"], height=480,
                    select_key="inv_tbl")
        st.caption(f"عدد الأصناف: {len(rows)}   |   قيمة المخزون: {money(sum(r[5] or 0 for r in rows))} ر.س")
        code = rows[sel[0]][0] if sel else None
        if open_once("item", code):
            item_dialog(code)
        elif code and st.button("🔎 عرض تفاصيل الصنف المحدد", key="item_reopen"):
            item_dialog(code)
    with t2:
        ilist = items()
        names = dict(ilist)
        code = st.selectbox("الصنف", [c for c, _ in ilist], index=None, placeholder="اختر الصنف...",
                            format_func=lambda c: f"{c} - {names[c]}", key="card_item")
        d1, d2 = date_range("card", start=date(date.today().year, 1, 1))
        if code:
            ok, res = run_safe(F.fetch_item_card, code, d1, d2)
            if ok:
                rows, final = res
                df = pd.DataFrame(rows, columns=["التاريخ", "الحركة", "الطرف", "الكمية", "الرصيد", "السعر", "القيمة", "ملاحظات"])
                table(df, height=420)
                final_text = f"الرصيد في نهاية الفترة: {final:,.3f}"
                st.success(final_text)
                download_on_demand("كارت الصنف PDF", f"card_{code}_{d1}_{d2}", F.build_item_card_pdf, ".pdf",
                                   "application/pdf", f"item_card_{code}.pdf", code, names[code], rows, final_text)


@st.dialog("تفاصيل الصنف", width="large")
def item_dialog(code):
    ok, d = run_safe(F.item_details, code)
    if not ok:
        return
    st.markdown(f"### {html.escape(d['name'])}")
    st.caption(f"الكود: {d['code']}   |   الوحدة: {d['unit']}   |   الباركود: {d['barcode'] or '-'}   |   "
               f"الكود القديم: {d['legacy'] or '-'}")
    m = st.columns(4)
    m[0].metric("الرصيد الحالي", f"{d['qty']:,.3f}")
    m[1].metric("متوسط التكلفة", money(d["cost"]))
    m[2].metric("قيمة المخزون", money(d["value"]))
    m[3].metric("سعر البيع", money(d["price"]) if d["price"] else "-",
                f"هامش {d['margin']:.1f}%" if d["margin"] is not None else None, delta_color="off")
    m = st.columns(3)
    m[0].metric("آخر تكلفة شراء", money(d["last_cost"]) if d["last_cost"] is not None else "-")
    m[1].metric("إجمالي المبيعات (كمية)", f"{d['sold_qty']:,.3f}")
    m[2].metric("إجمالي المشتريات (كمية)", f"{d['bought_qty']:,.3f}")
    st.markdown("**آخر الحركات**")
    table(pd.DataFrame(list(reversed(d["moves"])), columns=["التاريخ", "الحركة", "الطرف", "الكمية", "الرصيد", "السعر",
                                                           "القيمة", "ملاحظات"]), height=320)


# ============================================================
# كشوف الحساب
# ============================================================
def xlsx_bytes(df, sheet="Sheet1"):
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name=sheet[:31])
    return buf.getvalue()


def page_statement():
    header("📜 كشف الحساب الموحد")
    workspace("statement")


def statement_body():
    cust_ok, supp_ok = can_fin("customer"), can_fin("supplier")
    with F.db() as conn:
        rows = conn.execute("SELECT partner_code, name, is_customer, is_supplier FROM business_partners ORDER BY name").fetchall()
    allowed = [(c, n) for c, n, ic, is_ in rows if (ic and cust_ok) or (is_ and supp_ok)]
    if not allowed:
        st.warning("لا توجد صلاحية لعرض البيانات المالية للعملاء أو الموردين.")
        return
    scopes = [k for k in ("both", "customer", "supplier")
              if (k == "both" and cust_ok and supp_ok) or (k == "customer" and cust_ok) or (k == "supplier" and supp_ok)]
    if st.session_state.get("stmt_scope") not in scopes:
        st.session_state["stmt_scope"] = scopes[0]
    names = dict(allowed)
    if st.session_state.get("stmt_party") not in names:
        st.session_state.pop("stmt_party", None)
    c1, c2 = st.columns([2, 1])
    code = c1.selectbox("العميل / المورد (اختياري)", list(names), index=None,
                        placeholder="بدون اختيار ← ملخص أرصدة الفئة (سطر لكل طرف)",
                        format_func=lambda c: f"{c} - {names[c]}", key="stmt_party")
    scope = c2.selectbox("الفئة", scopes, format_func=F.SUMMARY_SCOPES.get, key="stmt_scope")
    c1, c2, c3 = st.columns([1, 1, 1])
    d_from = c1.date_input("من تاريخ", value=None, key="stmt_from", format="YYYY-MM-DD")
    d_to = c2.date_input("إلى تاريخ", value=None, key="stmt_to", format="YYYY-MM-DD")
    c3.caption("اترك التاريخين فارغين لعرض الرصيد اللحظي التراكمي من بداية العمل حتى الآن.")
    d_from = d_from.isoformat() if d_from else None
    d_to = d_to.isoformat() if d_to else None
    if d_from and d_to and d_from > d_to:
        st.error("تاريخ البداية بعد تاريخ النهاية.")
        return
    st.caption("📅 " + F.period_label(d_from, d_to))

    if not code:  # ------------------------------ العرض الإجمالي
        c1, c2 = st.columns([2, 1])
        q = c1.text_input("بحث في الملخص (الاسم / الكود / الهاتف)", key="stmt_q")
        hide_zero = c2.checkbox("إخفاء الأرصدة الصفرية", key="stmt_hide0")
        ok, res = run_safe(F.partners_balance_summary, scope, d_from, d_to, q, hide_zero)
        if not ok:
            return
        t = res["totals"]
        m = st.columns(4)
        m[0].metric("عدد الأطراف", len(res["rows"]))
        m[1].metric("مدين (لنا)", money(t["receivable"]))
        m[2].metric("دائن (علينا)", money(t["payable"]))
        m[3].metric("الصافي", money(abs(t["net"])), "مدين" if t["net"] > 0.005 else ("دائن" if t["net"] < -0.005 else "متوازن"),
                    delta_color="off")
        df = pd.DataFrame(F.summary_table_rows(res, bool(d_from)), columns=F.SUMMARY_HEADERS)
        if not d_from:
            df = df.drop(columns=["رصيد سابق"])
        sel = table(df, height=460, select_key="stmt_sum_tbl")
        if sel:
            r = res["rows"][sel[0]]
            st.button(f"📜 عرض كشف «{r['name']}» التفصيلي", type="primary", key="stmt_open_row",
                      on_click=set_state, kwargs={"stmt_party": r["code"], "stmt_sum_tbl": None})
        else:
            st.caption("💡 اختر سطراً لعرض الكشف التفصيلي للطرف.")
        c1, c2 = st.columns(2)
        with c1:
            download_on_demand("ملخص الأرصدة PDF", f"stmt_sum_{scope}_{d_from}_{d_to}_{q}_{hide_zero}",
                               F.build_balances_summary_pdf, ".pdf", "application/pdf", f"balances_{scope}.pdf",
                               scope, d_from, d_to, q, hide_zero)
        with c2:
            st.download_button("⬇️ تنزيل Excel", xlsx_bytes(df, "ملخص الأرصدة"), file_name=f"balances_{scope}.xlsx",
                               key="stmt_sum_xlsx")
        return

    # ------------------------------ الكشف التفصيلي لطرف محدد
    parties = F.statement_parties(code)
    if len(parties) > 1:
        st.info("🔗 كيان مزدوج مرتبط تلقائياً (بالاسم أو الرقم الضريبي) لإجراء المقاصة: " +
                "  +  ".join(f"{c} {n} ({r})" for c, n, r in parties))
    ok, res = run_safe(F.fetch_partner_statement, code, True, scope, True, d_from, d_to)
    if not ok:
        return
    info, srows, td, tc, bal, skeys = res
    c1, c2, c3 = st.columns(3)
    c1.metric("إجمالي المدين", f"{td:,.2f}")
    c2.metric("إجمالي الدائن", f"{tc:,.2f}")
    side = "مدين (عليه)" if bal > 0.005 else ("دائن (له)" if bal < -0.005 else "متوازن")
    c3.metric("الرصيد" + (" الصافي" if scope == "both" and len(parties) > 1 else ""), f"{abs(bal):,.2f}", side,
              delta_color="off")
    st.caption(f"الهاتف: {info[1] or '-'}   |   السجل التجاري: {info[2] or '-'}   |   الرقم الضريبي: {info[3] or '-'}")
    df = pd.DataFrame(srows, columns=["المرجع", "التاريخ", "النوع", "البيان", "مدين", "دائن", "الرصيد"])
    st.caption("🖱️ اضغط على أي فاتورة في الكشف لفتح تفاصيلها (الأصناف والكميات والضريبة) في تبويب مجاور "
               "وسدادها مباشرة، مع بقاء هذا الكشف مفتوحاً في تبويبه.")
    sel = table(df, height=460, select_key=f"stmt_det_{code}", cards=True)
    inv_rows = [i for i, k in enumerate(skeys) if k and k[0] in ("sales", "purchase")]
    pick = sel[0] if sel else None
    if open_once(f"stmt_det_{code}", pick) and pick is not None and skeys[pick] and skeys[pick][0] in ("sales", "purchase"):
        ws_open_tab("invoice", inv_view=tuple(skeys[pick][:2]))
        st.rerun()
    if inv_rows:
        c1, c2 = st.columns([3, 1])
        k_pick = c1.selectbox("فتح فاتورة من الكشف", inv_rows, index=None, key=f"stmt_inv_pick_{code}",
                              format_func=lambda i: f"{srows[i][2]} {srows[i][0]} ({srows[i][1]}) — {srows[i][4]} / {srows[i][5]}",
                              placeholder="اختر فاتورة لعرض تفاصيلها وسدادها", label_visibility="collapsed")
        if c2.button("عرض ↗", key=f"stmt_inv_go_{code}", disabled=k_pick is None, width="stretch"):
            ws_open_tab("invoice", inv_view=tuple(skeys[k_pick][:2]))
            st.rerun()
    c1, c2 = st.columns(2)
    with c1:
        download_on_demand("كشف الحساب PDF", f"stmt_{code}_{scope}_{d_from}_{d_to}", F.build_statement_pdf, ".pdf",
                           "application/pdf", f"statement_{code}.pdf", code, scope, d_from, d_to)
    with c2:
        st.download_button("⬇️ تنزيل Excel", xlsx_bytes(df, "كشف حساب"), file_name=f"statement_{code}.xlsx",
                           key="stmt_xlsx")
    st.button("↩️ العودة لملخص الأرصدة", key="stmt_back", on_click=set_state, kwargs={"stmt_party": None})


# ============================================================
# سجل المناديب
# ============================================================
def page_reps():
    header("🧑‍💼 سجل المناديب")
    st.caption("المناديب النشطون يظهرون كقائمة منسدلة في بطاقات العملاء والموردين وفي بيانات الشركة (المندوب الافتراضي).")
    q = st.text_input("بحث (الاسم / الهاتف / المنطقة / الكود)", key="reps_q")
    rows = F.list_reps(q=q)
    df = pd.DataFrame([(r[1], r[2], r[3] or "-", r[4] or "-", r[6] or "-", r[7] or "-", r[9], r[12], r[8] or "-",
                        "نشط" if r[10] else "موقوف") for r in rows],
                      columns=["الكود", "اسم المندوب", "الهاتف", "البريد", "المنطقة / خط السير", "المسمى", "العمولة %",
                               "عدد العملاء/الموردين", "تاريخ التعيين", "الحالة"])
    sel = table(df, height=320, select_key="reps_tbl")
    rec = rows[sel[0]] if sel else None
    st.markdown(f"#### {'✏️ تعديل المندوب ' + rec[1] if rec else '➕ مندوب جديد (الكود يُعطى تلقائياً)'}")
    with st.form(f"rep_form_{rec[0] if rec else 'new'}"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("اسم المندوب *", rec[2] if rec else "")
        phone = c2.text_input("رقم الهاتف", rec[3] if rec else "")
        email = c3.text_input("البريد الإلكتروني", rec[4] if rec else "")
        c1, c2, c3 = st.columns(3)
        nid = c1.text_input("رقم الهوية / الإقامة", rec[5] if rec else "")
        region = c2.text_input("المنطقة / خط السير", rec[6] if rec else "")
        job = c3.text_input("المسمى الوظيفي", rec[7] if rec else "مندوب مبيعات")
        c1, c2, c3 = st.columns(3)
        hire = c1.date_input("تاريخ التعيين", date.fromisoformat(rec[8]) if rec and rec[8] else None, format="YYYY-MM-DD")
        comm = c2.number_input("نسبة العمولة %", 0.0, 100.0, float(rec[9]) if rec else 0.0, step=0.5)
        active = c3.checkbox("نشط (يظهر في القوائم)", bool(rec[10]) if rec else True)
        notes = st.text_input("ملاحظات", rec[11] if rec else "")
        if rec:
            st.caption("تغيير الاسم ينتقل تلقائياً لكل الفواتير والأوامر وبطاقات الأطراف وسجل الفترات المرتبطة بالمندوب.")
        if st.form_submit_button("💾 حفظ", type="primary"):
            data = dict(name=name, phone=phone, email=email, national_id=nid, region=region, job_title=job,
                        hire_date=hire.isoformat() if hire else "", commission_pct=comm, active=active, notes=notes)
            ok, _ = run_safe(F.save_rep, rec[0] if rec else None, data, U()["username"])
            if ok:
                flash("تم حفظ بيانات المندوب.")
                st.session_state.pop("reps_tbl", None)
                st.rerun()
    if rec:
        c1, c2 = st.columns(2)
        if c1.button("🗑️ حذف / إيقاف المندوب", key=f"rep_del_{rec[0]}"):
            ok, res = run_safe(F.delete_rep, rec[0])
            if ok:
                flash("تم حذف المندوب." if res == "deleted" else
                      "المندوب مرتبط بمستندات سابقة، لذلك تم إيقافه بدلاً من حذفه.", "success" if res == "deleted" else "info")
                st.session_state.pop("reps_tbl", None)
                st.rerun()
        with F.db() as conn:
            ps = conn.execute("SELECT partner_code, name, is_customer, is_supplier, COALESCE(phone,'') FROM business_partners "
                              "WHERE rep_name=? ORDER BY name", (rec[2],)).fetchall()
        with c2.expander(f"👥 العملاء والموردون المسندون ({len(ps)})"):
            table(pd.DataFrame([(c, n, " و".join(x for x, f_ in (("عميل", ic), ("مورد", is_)) if f_) or "-", ph or "-")
                                for c, n, ic, is_, ph in ps], columns=["الكود", "الاسم", "الصفة", "الهاتف"]))


# ============================================================
# قائمة الدخل
# ============================================================
_INCOME_LINE = re.compile(r"^\s*(.*?)\s+(-?[\d,]+\.\d{2}) ر\.س\s*$")


def page_income():
    header("📈 قائمة الدخل وإقرار الضريبة")
    d1, d2 = date_range("income", start=date(date.today().year, 1, 1))
    ok, text = run_safe(F.build_income_report, d1, d2)
    if not ok:
        return
    _esc = __import__("html").escape  # أسماء الحسابات يكتبها المستخدم: تُهرَّب قبل عرضها كـ HTML
    html = ['<div class="fz-scroll"><table class="fz-income">']
    for line in text.splitlines():
        if not line.strip() or set(line.strip()) <= set("=-"):
            continue
        m = _INCOME_LINE.match(line)
        if m:
            label = m.group(1).strip()
            strong = any(w in label for w in ("صافي الربح", "مجمل الربح", "صافي المبيعات", "إجمالي المصروفات", "صافي الضريبة"))
            html.append(f'<tr class="{"h" if strong else ""}"><td>{_esc(label)}</td><td class="v">{m.group(2)} ر.س</td></tr>')
        else:
            html.append(f'<tr class="sec"><td colspan="2">{_esc(line.strip())}</td></tr>')
    html.append("</table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


# ============================================================
# التقارير
# ============================================================
def can_report(key):
    if not can(F.REPORT_GROUP_OF.get(key, "reports")):
        return False
    need = F.REPORT_FIN_REQUIRES.get(key)
    return need is None or can_fin("customer" if need.startswith("sales") else "supplier")


def page_reports():
    header("📊 التقارير وتحليل القرار")
    workspace("report")


def reports_body():
    catalog = [(g, [(k, t, d) for k, t, _, d in its if can_report(k)]) for g, its in F.REPORT_CATALOG]
    catalog = [(g, its) for g, its in catalog if its]
    if not catalog:
        st.warning("لا توجد تقارير مسموح بها.")
        return
    c1, c2 = st.columns(2)
    with c1:
        group = st.selectbox("المجموعة", [g for g, _ in catalog], key="rep_group")
    its = dict(catalog)[group]
    with c2:
        key = st.selectbox("التقرير", [k for k, _, _ in its], format_func=lambda k: F.REPORT_INDEX[k][0], key="rep_key")
    st.caption(F.REPORT_INDEX[key][2])
    d1, d2 = date_range("rep", start=date(date.today().year, 1, 1))
    with st.expander("خيارات إضافية"):
        c1, c2, c3, c4 = st.columns(4)
        grp = c1.selectbox("التجميع", ["شهري", "يومي"], key="rep_grp")
        stale = c2.number_input("أيام الركود", 0, 3650, F.STALE_DAYS_DEFAULT, key="rep_stale")
        late = c3.number_input("مهلة التأخر (يوم)", 0, 3650, 7, key="rep_late")
        expd = c4.number_input("تنبيه الصلاحية (يوم)", 0, 3650, 30, key="rep_exp")
    params = dict(d_from=d1, d_to=d2, group="month" if grp == "شهري" else "day", stale_days=int(stale),
                  late_days=int(late), expiry_days=int(expd))
    if key == "sales_mgr":
        c1, c2 = st.columns(2)
        reps = [""] + F.rep_names() + [F.NO_REP]
        params["rep"] = c1.selectbox("المندوب", reps, key="smr_rep", format_func=lambda x: x or "كل المناديب")
        custs = dict(partners("customer"))
        params["customer"] = c2.selectbox("العميل", [""] + list(custs), key="smr_cust",
                                          format_func=lambda c: f"{c} - {custs[c]}" if c else "كل العملاء")
    if d1 > d2:
        return
    ok, rep = run_safe(F.run_report, key, params)
    if not ok:
        return
    render_report(rep, f"{key}_{d1}_{d2}", key=key)


OPEN_INV_KEYS = ("open_inv_c", "open_inv_s")
DRILL_STATEMENT = {"rec": "customer", "age_c": "customer", "credit": "customer", "pay": "supplier", "age_s": "supplier"}


def _open_statement(code, scope):
    scopes_ok = {"customer": can_fin("customer"), "supplier": can_fin("supplier")}
    ws_open_tab("statement", stmt_party=code, stmt_scope=scope if scopes_ok.get(scope) else None,
                stmt_from=None, stmt_to=None, stmt_q=None, stmt_hide0=None)


def _drill(key, rep, idx):
    """تنفيذ الانتقال التفاعلي لسطر تقرير."""
    if idx is None or idx >= len(rep.rows):
        return
    row = rep.rows[idx]
    if key in DRILL_STATEMENT and row[0]:
        _open_statement(row[0], DRILL_STATEMENT[key])
        st.rerun()
    elif key in OPEN_INV_KEYS:
        ref = F.open_invoice_drill(key, row)
        if ref:
            ws_open_tab("invoice", inv_view=ref)
            st.rerun()
    elif key == "due_inv":
        kind = "sales" if str(row[0]).startswith("تحصيل") else "purchase"
        inv_id = F.find_invoice_id(kind, row[1])
        if inv_id:
            ws_open_tab("invoice", inv_view=(kind, inv_id))
            st.rerun()


def render_report(rep, dl_key, height=None, key=None):
    """عرض تقرير (جدول قابل للتمرير + رسم + ملاحظات + تصدير Excel/PDF عند الطلب). مشترك بين الصفحات."""
    st.markdown(f"### {rep.title}")
    if rep.subtitle:
        st.caption(rep.subtitle)
    cols = [h for h, _ in rep.columns]
    rows = list(rep.rows) + ([rep.total] if rep.total else [])
    df = pd.DataFrame([[F.fmt_cell(v, k) if k in ("t",) or isinstance(v, str) else v
                        for v, (_, k) in zip(r, rep.columns)] for r in rows], columns=cols)
    kinds = dict(rep.columns)
    for c, k in kinds.items():
        if k != "t":
            df[c] = pd.to_numeric(df[c], errors="coerce")
    drill = key in DRILL_STATEMENT or key == "due_inv" or key in OPEN_INV_KEYS
    sel_key = f"drill_{dl_key}" if drill else None
    if drill:
        st.caption("🖱️ اضغط على أي سطر لفتحه في تبويب مجاور: " + ("كشف حساب الطرف التفصيلي (ومنه الفواتير والسداد)."
                                               if key != "due_inv" else "تفاصيل الفاتورة وسدادها."))
    sel = table(df, money_cols=[c for c, k in kinds.items() if k == "m"], qty_cols=[c for c, k in kinds.items() if k == "q"],
                pct_cols=[c for c, k in kinds.items() if k == "p"], height=height or min(560, 60 + 35 * len(df)),
                select_key=sel_key, cards=True if drill else None, reverse=not getattr(rep, "ltr", False),
                highlight=[cols[i] for i in getattr(rep, "highlight", ())])
    if drill:
        idx = sel[0] if sel else None
        if open_once(sel_key, idx):
            _drill(key, rep, idx)
        # بديل مريح للمس (يظهر على كل الشاشات، وأساسي على الموبايل حيث يُعرض التقرير كبطاقات)
        opts = [i for i, r in enumerate(rep.rows) if r[0]]
        if key in DRILL_STATEMENT:
            label = lambda i: f"{rep.rows[i][1]} — {rep.rows[i][0]}"  # noqa: E731
        elif key in OPEN_INV_KEYS:
            label = lambda i: f"{rep.rows[i][0]} — {rep.rows[i][3]} (متبقي {money(rep.rows[i][8])})"  # noqa: E731
        else:
            label = lambda i: f"{rep.rows[i][0]} {rep.rows[i][1]} — {rep.rows[i][5]} ({money(rep.rows[i][6])})"  # noqa: E731
        c1, c2 = st.columns([3, 1])
        pick = c1.selectbox("فتح سريع", opts, index=None, format_func=label, key=f"drill_pick_{dl_key}",
                            placeholder="اختر طرفاً لفتح كشفه" if key != "due_inv" else "اختر فاتورة لعرضها وسدادها",
                            label_visibility="collapsed")
        if c2.button("فتح في تبويب ↗", key=f"drill_go_{dl_key}", disabled=pick is None, width="stretch"):
            _drill(key, rep, pick)
    if rep.chart and rep.rows:
        li, vi = rep.chart
        cdf = pd.DataFrame([(str(r[li]), r[vi]) for r in rep.rows[:15] if isinstance(r[vi], (int, float))],
                           columns=[cols[li], cols[vi]])
        if not cdf.empty:
            st.bar_chart(cdf.set_index(cols[li]), color=BLUE)
    for n in rep.notes:
        st.caption("• " + n)
    c1, c2 = st.columns(2)
    name = dl_key.split("_")[0]
    with c1:
        download_on_demand("Excel", f"rx_{dl_key}", F.export_report_xlsx, ".xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"report_{name}.xlsx", rep)
    with c2:
        download_on_demand("PDF", f"rp_{dl_key}", F.build_report_pdf, ".pdf", "application/pdf",
                           f"report_{name}.pdf", rep)


# ============================================================
# الحسابات: سجل الخصومات اللاحقة وأثرها
# ============================================================
def page_discounts():
    header("🏷️ الخصومات اللاحقة والإشعارات الدائنة")
    b1, b2, _ = st.columns([1, 1, 2])
    if can("purchase.discount"):
        b1.button("➕ تسجيل خصم لاحق من مورد", type="primary", width="stretch", key="dr_new_p",
                  on_click=go_discount, args=("purchase",))
    if can("sales.discount"):
        b2.button("➕ تسجيل إشعار دائن لعميل", type="primary", width="stretch", key="dr_new_s",
                  on_click=go_discount, args=("sales",))
    c1, c2 = st.columns([1, 2])
    with c1:
        typ = st.selectbox("النوع", ["الكل", "خصومات الموردين", "إشعارات العملاء"], key="dr_type")
    with c2:
        d1, d2 = date_range("dr", start=date(date.today().year, 1, 1))
    kind = {"خصومات الموردين": F.DISC_PURCHASE, "إشعارات العملاء": F.DISC_SALES}.get(typ)
    ok, data = run_safe(F.post_discount_register, kind, d1, d2)
    if not ok:
        return
    sup = [r for r in data if r["kind"] == F.DISC_PURCHASE]
    cus = [r for r in data if r["kind"] == F.DISC_SALES]
    m = st.columns(6)
    m[0].metric("عدد الإشعارات", f"{len(data)}")
    m[1].metric("خصومات الموردين", f"{money(sum(r['total'] for r in sup))}")
    m[2].metric("إشعارات العملاء", f"{money(sum(r['total'] for r in cus))}")
    m[3].metric("خفض المخزون", f"{money(sum(r['stock'] for r in data))}")
    m[4].metric("أثر الأرباح", f"{money(sum(r['pl'] for r in data))}")
    m[5].metric("أثر الضريبة", f"{money(sum(r['vat'] for r in sup) - sum(r['vat'] for r in cus))}")
    df = pd.DataFrame([(r["no"], r["date"], "مورد" if r["kind"] == F.DISC_PURCHASE else "عميل",
                        f"{r['party_code']} - {r['party']}", r["invoice"] or "-", r["method"], r["subtotal"], r["vat"],
                        r["total"], r["stock"], r["pl"], r["journal"] or "-") for r in data],
                      columns=["رقم الإشعار", "التاريخ", "النوع", "الطرف", "الفاتورة", "المعالجة", "قبل الضريبة", "الضريبة",
                               "الإجمالي", "خفض المخزون", "أثر الأرباح", "القيد"])
    table(df, money_cols=["قبل الضريبة", "الضريبة", "الإجمالي", "خفض المخزون", "أثر الأرباح"], height=min(460, 60 + 35 * len(df)))
    st.caption("أثر الأرباح: خصم المورد المباشر وما يخص بضاعة مباعة (+ خصم مكتسب)، والإشعار الدائن للعميل (- يخفض صافي "
               "المبيعات). خفض المخزون: الجزء الذي خفّض متوسط تكلفة الرصيد المتبقي ويظهر لاحقاً في تكلفة المبيعات.")
    if data:
        ids = {r["id"]: r for r in data}
        pick = st.selectbox("تفاصيل إشعار", list(ids), format_func=lambda v: f"{ids[v]['no']} - {ids[v]['party']} "
                            f"({money(ids[v]['total'])})", key="dr_pick")
        if pick:
            r = ids[pick]
            d, lines = F.load_post_discount(pick)
            st.markdown(f'<div class="fz-card"><h4>{r["no"]} • {html.escape(r["party"])}</h4>'
                        f'<span class="fz-chip">{r["balance_effect"]}: {money(r["total"])}</span>'
                        f'<span class="fz-chip">{r["vat_effect"]}: {money(r["vat"])}</span>'
                        f'<span class="fz-chip">خفض المخزون: {money(r["stock"])}</span>'
                        f'<span class="fz-chip">أثر الأرباح: {money(r["pl"])}</span>'
                        f'<span class="fz-chip">القيد الآلي: {r["journal"] or "غير مرحّل"}</span></div>',
                        unsafe_allow_html=True)
            if lines:
                table(pd.DataFrame([(c, n, a, ap, a - ap) for c, n, a, ap in lines],
                                   columns=["الكود", "الصنف", "الخصم الموزع", "خُفّض من المخزون", "خصم مكتسب (بضاعة مباعة)"]),
                      money_cols=["الخصم الموزع", "خُفّض من المخزون", "خصم مكتسب (بضاعة مباعة)"])
            if r["journal"]:
                with F.db() as conn:
                    jl = conn.execute("SELECT account_code, account_name, debit, credit, COALESCE(memo,'') FROM journal_lines "
                                      "WHERE entry_id=? ORDER BY line_id", (r["journal"],)).fetchall()
                st.markdown("**القيد اليومي الآلي**")
                table(pd.DataFrame(jl, columns=["الحساب", "اسم الحساب", "مدين", "دائن", "البيان"]), money_cols=["مدين", "دائن"])
            download_on_demand("الإشعار PDF", f"drp_{pick}", F.build_post_discount_pdf, ".pdf", "application/pdf",
                               f"credit_note_{r['no']}.pdf", pick)
        rep = F.run_report("pdisc", dict(d_from=d1, d_to=d2))
        if kind:
            rep.rows = [x for x in rep.rows if x[2] == ("مورد" if kind == F.DISC_PURCHASE else "عميل")]
        download_on_demand("السجل Excel", f"drx_{typ}_{d1}_{d2}", F.export_report_xlsx, ".xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"post_discounts_{d1}_{d2}.xlsx",
                           rep)


# ============================================================
# القيود اليومية ودليل الحسابات
# ============================================================
def page_journal():
    header("📓 القيود اليومية")
    d1, d2 = date_range("jv", default_days=90)
    att = F.attachment_counts("journal")
    with F.db() as conn:
        heads = conn.execute(
            "SELECT e.entry_id, e.entry_date, COALESCE(e.description,''), COALESCE(SUM(l.debit),0), "
            "CASE WHEN e.source_type IS NULL THEN 'يدوي' ELSE 'آلي' END FROM journal_entries e "
            "LEFT JOIN journal_lines l ON l.entry_id=e.entry_id WHERE e.entry_date BETWEEN ? AND ? "
            "GROUP BY e.entry_id ORDER BY e.entry_date DESC, e.entry_id DESC", (d1, d2)).fetchall()
    st.caption("🖱️ اختر قيداً لعرض بنوده ومرفقاته.")
    sel = table(pd.DataFrame([h + (f"📎 {att[h[0]]}" if att.get(h[0]) else "-",) for h in heads],
                             columns=["رقم القيد", "التاريخ", "البيان", "الإجمالي", "النوع", "المرفقات"]),
                money_cols=["الإجمالي"], select_key="jv_tbl")
    if sel:
        eid = heads[sel[0]][0]
        with F.db() as conn:
            lines = conn.execute("SELECT account_code, account_name, debit, credit, COALESCE(memo,'') FROM journal_lines "
                                 "WHERE entry_id=? ORDER BY line_id", (eid,)).fetchall()
        st.markdown(f"#### بنود القيد {eid}")
        table(pd.DataFrame(lines, columns=["الحساب", "اسم الحساب", "مدين", "دائن", "البيان"]), money_cols=["مدين", "دائن"])
        attachments_box("journal", eid)
    st.caption("القيود الآلية (مثل قيود الخصم اللاحق) تُعدَّل بتعديل مستندها. القيود اليدوية تُسجَّل من برنامج سطح المكتب.")


def page_coa():
    header("🗂️ دليل الحسابات (شجرة هرمية)")
    t1, t2, t3 = st.tabs(["🌳 الشجرة", "🌿 حساب فرعي جديد", "📒 كشف حساب تحليلي"])
    tree = F.account_tree_rows()
    with t1:
        q = st.text_input("بحث بالكود أو الاسم", key="coa_q")
        rows = [r for r in tree if not q or F.search_norm(q) in F.search_norm(f"{r[0]} {r[1]}")]
        table(pd.DataFrame([(c, ("\u2003" * lvl) + ("▸ " if g else "") + n, t, f"{'تجميعي' if g else 'فرعي'} ({lvl + 1})")
                            for c, n, t, lvl, g in rows], columns=["الكود", "الحساب", "النوع", "المستوى"]), height=560)
    names = {c: n for c, n, *_ in tree}
    with t2:
        if not (U()["is_admin"] or can("modify")):
            st.info("إنشاء الحسابات الفرعية متاح للمدير ولمن لديه صلاحية التعديل.")
        else:
            st.caption("أمثلة: «سلفة الموظف محمد» تحت 11410 ذمم العاملين، «خزينة فرع جدة» تحت 11110، «حساب جاري» تحت 11120.")
            parent = st.selectbox("الحساب الرئيسي", list(names), index=None, placeholder="اختر الحساب الأب...",
                                  format_func=lambda c: f"{c} - {names[c]}", key="sub_parent")
            if parent:
                with F.db() as conn:
                    st.caption(f"الكود التلقائي للحساب الجديد: {F.next_child_account_code(conn, parent)}")
            sub_name = st.text_input("اسم الحساب الفرعي", key="sub_name")
            is_group = st.checkbox("حساب تجميعي (ستُضاف تحته فروع أخرى)", key="sub_group")
            if st.button("💾 إنشاء الحساب الفرعي", type="primary", disabled=not parent, key="sub_go"):
                ok, code = run_safe(F.create_sub_account, parent, sub_name, None, is_group)
                if ok:
                    flash(f"تم إنشاء الحساب {code} - {sub_name}")
                    st.session_state.pop("sub_name", None)
                    st.rerun()
    with t3:
        acc = st.selectbox("الحساب (الرئيسي يجمع كل فروعه)", list(names), index=None, placeholder="اختر الحساب...",
                           format_func=lambda c: f"{c} - {names[c]}", key="acc_st")
        d1, d2 = date_range("acc_st", start=date(date.today().year, 1, 1))
        if acc:
            ok, s_ = run_safe(F.account_statement, acc, d1, d2)
            if ok:
                if s_["group"]:
                    st.caption(f"يشمل {s_['accounts'] - 1} حساباً فرعياً تحت {s_['code']} - {s_['name']}")
                m = st.columns(4)
                m[0].metric("رصيد أول المدة", money(s_["opening"]))
                m[1].metric("المسحوب / مدين", money(s_["debit"]))
                m[2].metric("المسدد / دائن", money(s_["credit"]))
                m[3].metric("المتبقي (الرصيد)", money(s_["closing"]))
                df = pd.DataFrame(s_["rows"], columns=["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"])
                table(df, money_cols=["مدين", "دائن", "الرصيد"], height=420)
                rep = F.Report(f"كشف حساب تحليلي: {s_['code']} - {s_['name']}",
                               [("التاريخ", "t"), ("المرجع", "t"), ("الحساب", "t"), ("البيان", "t"), ("مدين", "m"),
                                ("دائن", "m"), ("الرصيد", "m")], s_["rows"],
                               ("", "", "", "الإجمالي", s_["debit"], s_["credit"], s_["closing"]),
                               [f"رصيد أول المدة: {money(s_['opening'])}"], f"من {d1} إلى {d2}")
                c1, c2 = st.columns(2)
                with c1:
                    download_on_demand("Excel", f"accx_{acc}_{d1}_{d2}", F.export_report_xlsx, ".xlsx",
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       f"account_{acc}.xlsx", rep)
                with c2:
                    download_on_demand("PDF", f"accp_{acc}_{d1}_{d2}", F.build_report_pdf, ".pdf", "application/pdf",
                                       f"account_{acc}.pdf", rep)


# ============================================================
# الشركاء والرفع الجماعي (Excel / CSV)
# ============================================================
def _read_upload(up):
    suffix = os.path.splitext(up.name)[1].lower()
    if suffix not in (".xlsx", ".xls", ".csv"):
        raise F.UserError("الملف يجب أن يكون Excel (xlsx) أو CSV.")
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    try:
        with open(path, "wb") as f:
            f.write(up.getvalue())
        return F.read_table_file(path)
    finally:
        os.remove(path)


def page_partners(role):
    sales = role == "customer"
    header("👥 العملاء" if sales else "🏭 الموردون")
    who = "العميل" if sales else "المورد"
    q = st.text_input("بحث (الاسم / الكود / الهاتف / الرقم الضريبي / المندوب)", key=f"pp_q_{role}")
    rows = F.list_partners(role, q)
    fin = can_fin(role)
    df = pd.DataFrame([(r[0], r[1], r[2] or "-", r[3] or "-", r[4] or "-",
                        (f"{abs(r[5]):,.2f} {'مدين' if r[5] > 0 else 'دائن'}" if abs(r[5]) > 0.005 else "-") if fin else "🔒",
                        r[-1] or "-", r[6] or "-") for r in rows],
                      columns=["الكود", "الاسم", "الهاتف", "الرقم الضريبي", "المندوب", "الرصيد الافتتاحي", "مرتبط بـ",
                               "الكود القديم"])
    st.caption(f"العدد: {len(rows)}   |   «مرتبط بـ» = نفس الكيان في الشجرة الأخرى (يُربط تلقائياً بالاسم أو الرقم الضريبي "
               "لإجراء المقاصة في الكشف الموحد).   🖱️ اختر صفاً للتعديل.")
    sel = table(df, height=380, select_key=f"pp_tbl_{role}")
    rec = rows[sel[0]] if sel else None
    code = rec[0] if rec else None
    st.markdown(f"#### {'✏️ تعديل ' + code if code else '➕ ' + who + ' جديد (الكود يُولَّد تلقائياً)'}")
    if not code:
        with F.db() as conn:
            st.caption(f"الكود التالي: {F.peek_serial(conn, role)[0]}")
    sfx = f"{role}_{code or 'new'}"
    with st.form(f"pp_form_{sfx}"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input(f"اسم {who}", rec[1] if rec else "")
        phone = c2.text_input("الهاتف", rec[2] if rec else "")
        email = c3.text_input("البريد", rec[9] if rec else "")
        c1, c2, c3 = st.columns(3)
        address = c1.text_input("العنوان", rec[7] if rec else "")
        cr = c2.text_input("السجل التجاري", rec[8] if rec else "")
        vat = c3.text_input("الرقم الضريبي", rec[3] if rec else "")
        c1, c2 = st.columns(2)
        rep_opts = [""] + F.rep_names(include=(rec[4] if rec and rec[4] else None))
        rep = c1.selectbox("المندوب (افتراضي) — من سجل المناديب", rep_opts,
                           index=rep_opts.index(rec[4]) if rec and rec[4] in rep_opts else 0,
                           format_func=lambda x: x or "— بدون مندوب —",
                           help="تغيير المندوب يسري من التاريخ المحدد ولا يغيّر حركات الفترات السابقة "
                                "(سجل تاريخي لمنع تداخل العمولات). لإضافة مندوب جديد: صفحة «سجل المناديب».")
        rep_date = c2.date_input("تاريخ سريان المندوب", date.today(), format="YYYY-MM-DD")
        opening = None
        if U()["is_admin"] or can("modify"):
            with F.db() as conn:
                ob = conn.execute("SELECT COALESCE(opening_debit,0), COALESCE(opening_credit,0), COALESCE(opening_date,'') "
                                  "FROM business_partners WHERE partner_code=?", (code,)).fetchone() if code else (0, 0, "")
            c1, c2, c3 = st.columns(3)
            od = c1.number_input("رصيد افتتاحي مدين (عليه)", min_value=0.0, value=float(ob[0]), format="%.2f")
            oc = c2.number_input("رصيد افتتاحي دائن (له)", min_value=0.0, value=float(ob[1]), format="%.2f")
            odate = c3.text_input("تاريخ الرصيد (YYYY-MM-DD)", ob[2] or "")
            opening = (od, oc, odate)
        with F.db() as conn:
            terms = conn.execute("SELECT COALESCE(credit_limit,0), credit_days, COALESCE(credit_policy,'') "
                                 "FROM business_partners WHERE partner_code=?", (code,)).fetchone() if code else (0, None, "")
        st.markdown("**💳 الائتمان وشروط السداد**")
        c1, c2, c3 = st.columns(3)
        days = c2.number_input("مدة السداد بالأيام (0 = الافتراضية)", min_value=0, max_value=3650,
                               value=int(terms[1] or 0), help=f"الافتراضية {F.default_credit_days()} يوماً. "
                               "تحدد تاريخ استحقاق الفواتير في توقع التدفق النقدي.")
        credit = (None, int(days) if days else "", None)
        if sales:
            limit = c1.number_input("الحد الائتماني (0 = بدون حد)", min_value=0.0, value=float(terms[0]), step=1000.0,
                                    format="%.2f")
            pol_opts = [""] + list(F.CREDIT_POLICIES)
            pol = c3.selectbox("سياسة التجاوز", pol_opts, index=pol_opts.index(terms[2]) if terms[2] in pol_opts else 0,
                               format_func=lambda k: F.CREDIT_POLICIES.get(k, f"الافتراضية ({F.CREDIT_POLICIES[F.default_policy()]})"))
            credit = (limit, credit[1], pol)
        if st.form_submit_button("💾 حفظ", type="primary"):
            ok, new_code = run_safe(F.save_partner_record, role, code, name, phone, email, address, cr, vat, rep,
                                    rep_date.isoformat(), opening, U()["username"], credit)
            if ok:
                links = [c for c in F.linked_partner_codes(new_code) if c != new_code]
                flash(f"تم الحفظ بالكود {new_code}" + (f" — مرتبط تلقائياً بـ {' ، '.join(links)}" if links else ""))
                st.session_state.pop(f"pp_tbl_{role}", None)
                st.rerun()
    if code and sales and can_fin("customer"):
        cst = F.credit_status(code, 0.0, "order")
        if cst["limit"] > 0.005:
            m = st.columns(4)
            m[0].metric("الحد الائتماني", money(cst["limit"]))
            m[1].metric("إجمالي التعرض", money(cst["exposure"]))
            m[2].metric("المتاح", money(cst["available"]))
            m[3].metric("نسبة الاستخدام", f"{cst['usage']:,.1f}%")
            if cst["exceeded"]:
                st.error("⛔ العميل متجاوز لحده الائتماني.")
    if code:
        c1, c2, c3 = st.columns(3)
        c1.button("📜 الكشف الموحد", key=f"pp_st_{code}", on_click=go_to, args=("statement",),
                  kwargs={"stmt_party": code, "stmt_scope": "both"})
        hist = F.partner_rep_history(code)
        with st.expander(f"🕘 سجل المناديب ({len(hist)})"):
            table(pd.DataFrame([(r, s_, e_ or "حتى الآن", by or "-") for r, s_, e_, by, _at in hist],
                               columns=["المندوب", "من", "إلى", "بواسطة"]))
        if U()["is_admin"] or can("modify"):
            if c3.checkbox("تأكيد الحذف", key=f"pp_delc_{code}") and c3.button("🗑️ حذف", key=f"pp_del_{code}"):
                ok, _ = run_safe(F.delete_partner, code)
                if ok:
                    flash(f"تم حذف {code}.")
                    st.session_state.pop(f"pp_tbl_{role}", None)
                    st.rerun()


# ============================================================
# مركز استيراد البيانات (Excel / CSV)
# ============================================================
IMPORT_ICONS = {"items": "📦", "customers": "👥", "suppliers": "🏭", "cash": "🏦"}


def page_import():
    header("📥 مركز استيراد البيانات والترحيل")
    st.markdown(
        '<div class="fz-steps">'
        '<div class="fz-step"><b>1</b>نزّل النموذج (Excel أو CSV) واملأ بياناتك.</div>'
        '<div class="fz-step"><b>2</b>ارفع الملف لمعاينة النتيجة كاملة <u>بدون حفظ</u>.</div>'
        '<div class="fz-step"><b>3</b>نفّذ الاستيراد؛ تُحفظ نسخة احتياطية تلقائياً قبل التنفيذ.</div>'
        '</div><div style="margin:10px 0 4px">'
        '<span class="fz-chip">الأصناف: ELH-0001</span><span class="fz-chip">العملاء: C0001</span>'
        '<span class="fz-chip">الموردون: S0001</span><span class="fz-chip">الكود القديم يُحفظ كمرجع</span></div>',
        unsafe_allow_html=True)
    kinds = ["items", "customers", "suppliers", "cash"]
    for tab, kind in zip(st.tabs([f"{IMPORT_ICONS[k]} {F.IMPORT_KINDS[k]['title']}" for k in kinds]), kinds):
        with tab:
            if kind == "cash":
                cash_quick_form()
            import_kind_ui(kind)


def cash_quick_form():
    st.markdown("#### ⚡ إعداد سريع لأرصدة الخزينة والبنوك")
    with F.db() as conn:
        names = F.cash_account_names(conn)
        codes = F.cash_accounts(conn)
    as_of, amounts = F.get_cash_opening()
    with st.form("cash_opening_form"):
        d = st.date_input("تاريخ الأرصدة", date.fromisoformat(as_of) if as_of else date.today(), format="YYYY-MM-DD")
        cols = st.columns(3)
        values = {}
        for i, name in enumerate(names):
            values[name] = cols[i % 3].number_input(f"{name} ({codes.get(name, '-')})", value=float(amounts.get(name, 0.0)),
                                                    step=100.0, format="%.2f", key=f"cash_open_{name}")
        st.caption(f"الإجمالي الحالي بالنموذج: {money(sum(values.values()))} ر.س — يستبدل قيد الأرصدة الافتتاحية السابق.")
        if st.form_submit_button("💾 حفظ أرصدة الخزينة والبنوك", type="primary"):
            ok, _ = run_safe(F.set_cash_opening, d.isoformat(), values)
            if ok:
                flash("تم حفظ الأرصدة الافتتاحية للخزينة والبنوك.")
                st.rerun()
    st.divider()


def import_kind_ui(kind):
    spec = F.IMPORT_KINDS[kind]
    chips = "".join(f'<span class="fz-chip">{html.escape(c)}</span>' for c in spec["columns"])
    st.markdown(f'<div class="fz-card"><h4>{IMPORT_ICONS[kind]} {spec["title"]} من ملف</h4>'
                f'<div style="color:#5b6b82">{html.escape(spec["help"])}</div><div style="margin-top:8px">{chips}</div></div>',
                unsafe_allow_html=True)
    c1, c2, _ = st.columns([1, 1, 2])
    c1.download_button("⬇️ نموذج Excel", build_file(F.write_import_template, ".xlsx", kind),
                       file_name=f"template_{kind}.xlsx", key=f"tplx_{kind}", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇️ نموذج CSV", build_file(F.write_import_template, ".csv", kind), file_name=f"template_{kind}.csv",
                       key=f"tplc_{kind}", width="stretch", mime="text/csv")
    opts = {}
    if kind == "items":
        a, b = st.columns(2)
        opts["update_existing"] = a.checkbox("تحديث بيانات الأصناف المطابقة (الوحدة، سعر البيع، الباركود، الرصيد الافتتاحي)",
                                             True, key="imp_upd")
        opts["recode_matched"] = b.checkbox("توحيد أكواد الأصناف المطابقة إلى ELH-XXXX (يُحفظ الكود القديم كمرجع)",
                                            False, key="imp_recode")
    else:
        opts["as_of"] = st.date_input("تاريخ الرصيد الافتتاحي", date.today(), key=f"imp_asof_{kind}",
                                      format="YYYY-MM-DD").isoformat()
    up = st.file_uploader("اسحب الملف هنا أو اختره (Excel / CSV)", type=["xlsx", "xls", "csv"], key=f"imp_up_{kind}")
    if not up:
        return
    ok, df = run_safe(_read_upload, up)
    if not ok:
        return
    ok, res = run_safe(F.run_import, kind, df, True, **opts)
    if not ok:
        return
    _import_result(kind, res, preview=True)
    if st.button("✅ تنفيذ الاستيراد والحفظ", type="primary", key=f"imp_go_{kind}"):
        F.backup_database("before_import")
        ok, res = run_safe(F.run_import, kind, df, False, **opts)
        if ok:
            flash("✅ تم الاستيراد: " + F.import_stats_text(res["stats"]))
            st.session_state.pop(f"imp_up_{kind}", None)
            st.rerun()


def _import_result(kind, res, preview):
    stx = res["stats"]
    st.markdown("#### " + ("🔎 معاينة (لم يُحفظ شيء بعد)" if preview else "✅ النتيجة"))
    keys = [k for k in ("rows", "added", "updated", "recoded", "opening_set", "accounts", "duplicates", "bad")
            if k in stx and (stx[k] or k in ("rows", "added"))]
    cols = st.columns(max(len(keys), 1))
    for c, k in zip(cols, keys):
        c.metric(F.IMPORT_STAT_LABELS[k], f"{stx[k]:,}")
    money_keys = [k for k in ("total_debit", "total_credit", "stock_value", "total") if k in stx]
    if money_keys:
        st.caption("   |   ".join(f"{F.IMPORT_STAT_LABELS[k]}: {money(stx[k])} ر.س" for k in money_keys))
    if res["problems"]:
        with st.expander(f"⚠️ ملاحظات ({len(res['problems'])})", expanded=True):
            for p in res["problems"][:200]:
                st.write("• " + p)
    if res["preview"]:
        cols_ = F.import_preview_columns(kind)
        dfp = pd.DataFrame(res["preview"], columns=[c or "-" for c in cols_])
        if "-" in dfp.columns:
            dfp = dfp.drop(columns=["-"])
        table(dfp, money_cols=["مدين (عليه)", "دائن (له)", "الرصيد"], height=min(420, 60 + 35 * len(dfp)))


# ============================================================
# المستخدمون والصلاحيات (للمدير فقط)
# ============================================================
def page_users():
    header("👤 المستخدمون والصلاحيات")
    if not U()["is_admin"]:
        st.error("هذه الصفحة لمدير النظام فقط.")
        return
    users = F.list_users()
    df = pd.DataFrame([(u[1], u[2], "مدير" if u[3] else "مستخدم", "نشط" if u[4] else "موقوف", u[6] or "-",
                        "مقفل مؤقتاً" if u[8] else "") for u in users],
                      columns=["اسم المستخدم", "الاسم", "النوع", "الحالة", "آخر دخول", "قفل"])
    table(df)
    ids = {u[0]: u for u in users}
    uid = st.selectbox("المستخدم", [None] + list(ids), format_func=lambda v: "➕ مستخدم جديد" if v is None else
                       f"✏️ {ids[v][1]} - {ids[v][2]}", key="usr_pick")
    rec = ids.get(uid)
    sfx = uid or "new"
    perms_now = set((rec[5] or "").split(",")) if rec else set()
    c1, c2 = st.columns(2)
    uname = c1.text_input("اسم المستخدم (إنجليزي)", rec[1] if rec else "", key=f"usr_u_{sfx}")
    full = c2.text_input("الاسم الكامل", rec[2] if rec else "", key=f"usr_f_{sfx}")
    c1, c2 = st.columns(2)
    pw = c1.text_input("كلمة المرور" + (" (اتركها فارغة لعدم التغيير)" if rec else ""), type="password", key=f"usr_p_{sfx}")
    pw2 = c2.text_input("تأكيد كلمة المرور", type="password", key=f"usr_p2_{sfx}")
    c1, c2 = st.columns(2)
    is_admin = c1.checkbox("مدير النظام (كل الصلاحيات)", bool(rec[3]) if rec else False, key=f"usr_a_{sfx}")
    active = c2.checkbox("نشط", bool(rec[4]) if rec else True, key=f"usr_act_{sfx}")
    chosen = []
    if not is_admin:
        st.markdown("**الصلاحيات**")
        for key, label in F.PERMISSIONS:
            on = st.checkbox(label, key in perms_now, key=f"usr_perm_{sfx}_{key}")
            if on:
                chosen.append(key)
                if key in F.SUB_PERMISSIONS:
                    cols = st.columns(3)
                    for i, (sk, sl) in enumerate(F.SUB_PERMISSIONS[key]):
                        default = sk in perms_now or not any(s in perms_now for s, _ in F.SUB_PERMISSIONS[key])
                        if cols[i % 3].checkbox(sl, default, key=f"usr_perm_{sfx}_{sk}"):
                            chosen.append(sk)
    b1, b2, b3 = st.columns(3)
    if b1.button("💾 حفظ المستخدم", type="primary", key=f"usr_save_{sfx}"):
        ok, msg = run_safe(F.save_user_account, uid, uname, full, pw, pw2, is_admin, active, chosen, U()["user_id"])
        if ok:
            flash(msg + " تسري الصلاحيات من الدخول التالي للمستخدم.")
            st.rerun()
    if rec and rec[8] and b2.button("🔓 فك القفل", key=f"usr_unlock_{sfx}"):
        with F.db() as conn:
            conn.execute("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE user_id=?", (uid,))
        flash("تم فك قفل المستخدم.")
        st.rerun()
    if rec and rec[1] != F.DEFAULT_ADMIN_USER:
        confirm = b3.checkbox("تأكيد الحذف", key=f"usr_delc_{sfx}")
        if b3.button("🗑️ حذف المستخدم", disabled=not confirm, key=f"usr_del_{sfx}"):
            ok, name = run_safe(F.delete_user_account, uid, U()["user_id"])
            if ok:
                flash(f"تم حذف المستخدم {name}.")
                st.session_state.pop("usr_pick", None)
                st.rerun()


# ============================================================
# الإعدادات: الترقيم التسلسلي، حسابات قيد الخصم، النسخ الاحتياطي
# ============================================================
REQUIRED_TABLES = {"business_partners", "inventory", "sales_invoices", "purchase_invoices", "chart_of_accounts"}


def _restore_db(data):
    if not data.startswith(b"SQLite format 3\x00"):
        raise F.UserError("الملف ليس قاعدة بيانات SQLite صالحة.")
    fd, tmp = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        with open(tmp, "wb") as f:
            f.write(data)
        conn = sqlite3.connect(tmp)
        try:
            if conn.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                raise F.UserError("قاعدة البيانات المرفوعة تالفة.")
            tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        finally:
            conn.close()
        missing = REQUIRED_TABLES - tables
        if missing:
            raise F.UserError("الملف ليس قاعدة بيانات نظام الحاجري المالي (جداول ناقصة: " + "، ".join(sorted(missing)) + ").")
        F.backup_database("before_restore")
        src, dst = sqlite3.connect(tmp), sqlite3.connect(F.DB_FILE)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
    finally:
        os.remove(tmp)
    F.init_db()  # ترقية الهيكل للإصدار الحالي وفرض بيانات دخول المدير الرئيسي
    with F.db() as conn:
        F.audit_event(conn, F.ACT_EVENT, "استعادة قاعدة البيانات", "ملف مرفوع", new={"size": len(data)})


def company_form():
    st.caption("هذه البيانات القانونية تظهر على الفواتير الضريبية والإشعارات المطبوعة فقط؛ هوية البرنامج «نظام الحاجري المالي» منفصلة.")
    info = F.company_info()
    labels = dict(company_name="اسم المنشأة (عربي)", company_name_en="Company Name (EN)", vat_number="الرقم الضريبي",
                  cr_number="السجل التجاري", address="العنوان (عربي)", address_en="Address (EN)", bank_name="البنك",
                  iban="رقم الآيبان IBAN", phone="الهاتف", rep_name="اسم المندوب الافتراضي")
    with st.form("company_form"):
        vals = {}
        keys = list(labels)
        for i in range(0, len(keys), 2):
            cols = st.columns(2)
            for c, k in zip(cols, keys[i:i + 2]):
                if k == "rep_name":
                    opts = [""] + F.rep_names(include=info.get(k) or None)
                    vals[k] = c.selectbox(labels[k], opts, index=opts.index(info.get(k) or ""), key=f"comp_{k}",
                                          format_func=lambda x: x or "— بدون —")
                else:
                    vals[k] = c.text_input(labels[k], info.get(k) or "", key=f"comp_{k}")
        if st.form_submit_button("💾 حفظ بيانات المنشأة", type="primary"):
            ok, _ = run_safe(F.save_company_info, vals)
            if ok:
                flash("تم حفظ بيانات المنشأة المصدرة للفواتير.")
                st.rerun()


def page_settings():
    header("⚙️ بيانات الشركة والإعدادات")
    with F.db() as conn:
        comp = conn.execute("SELECT company_name, cr_number, vat_number, phone, address FROM company_profile WHERE id=1").fetchone()
        seqs = conn.execute("SELECT name, label, prefix, next_value, width FROM sequences ORDER BY rowid").fetchall()
    if comp:
        st.markdown(f"**{comp[0]}**  \nالسجل التجاري: {comp[1]}   |   الرقم الضريبي: {comp[2]}   |   الهاتف: {comp[3]}  \n{comp[4]}")
    if not U()["is_admin"]:
        return
    t0, t1, t2, t3, t4 = st.tabs(["🏢 المنشأة المصدرة للفواتير", "🔢 الترقيم التسلسلي", "🏷️ حسابات قيد الخصم اللاحق",
                                  "💾 قاعدة البيانات", "🩺 فحص سلامة البيانات"])
    with t4:
        integrity_tab()
    with t0:
        company_form()
    with t1:
        st.caption("كل تسلسل مستقل (فواتير المبيعات، المشتريات، إشعارات الخصم، العملاء، الموردين، الأصناف).")
        for name, label, prefix, nxt, width in seqs:
            c1, c2, c3, c4, c5 = st.columns([3, 2, 2, 2, 2])
            c1.markdown(f"**{label}**")
            p = c2.text_input("البادئة", prefix, key=f"seq_p_{name}")
            n = c3.number_input("الرقم التالي", 1, 10**9, int(nxt), key=f"seq_n_{name}")
            w = c4.number_input("عدد الأرقام", 1, 12, int(width), key=f"seq_w_{name}")
            if c5.button("حفظ", key=f"seq_s_{name}"):
                ok, _ = run_safe(F.update_sequence, name, p, int(n), int(w))
                if ok:
                    flash(f"تم تحديث تسلسل {label}.")
                    st.rerun()
    with t2:
        with F.db() as conn:
            cur = F.disc_accounts(conn)
            accs = conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE COALESCE(is_group,0)=0 "
                                "ORDER BY account_code").fetchall()
        anames = dict(accs)
        codes = list(anames)
        chosen = {}
        for key, (_, label) in F.DISC_ACCOUNT_DEFAULTS.items():
            chosen[key] = st.selectbox(label, codes, index=codes.index(cur[key]) if cur[key] in codes else None,
                                       format_func=lambda c: f"{c} - {anames[c]}", key=f"dacc_{key}")
        if st.button("💾 حفظ وإعادة ترحيل قيود الخصم", type="primary", key="dacc_save"):
            if any(v is None for v in chosen.values()):
                st.error("اختر كل الحسابات.")
            else:
                ok, _ = run_safe(F.set_disc_accounts, chosen)
                if ok:
                    flash("تم حفظ الحسابات وإعادة بناء القيود الآلية.")
                    st.rerun()
    with t3:
        if ON_STREAMLIT_CLOUD:
            st.warning("تنبيه: ملفات Streamlit Cloud مؤقتة؛ قد تُفقد البيانات عند إعادة تشغيل التطبيق أو تحديثه. "
                       "نزّل نسخة احتياطية بانتظام، واستعدها من هنا عند الحاجة.")
        st.caption(f"قاعدة البيانات النشطة: {os.path.basename(F.DB_FILE)}")
        download_on_demand("نسخة احتياطية من قاعدة البيانات", "dbbackup",
                           lambda path: _copy_db(path), ".db", "application/octet-stream",
                           f"elhagry_backup_{date.today():%Y%m%d}.db")
        up = st.file_uploader("استعادة قاعدة بيانات (ملف ‎.db)", type=["db", "sqlite", "sqlite3"], key="db_restore")
        if up:
            confirm = st.checkbox("أؤكد استبدال البيانات الحالية بالملف المرفوع (تُحفظ نسخة من الحالية تلقائياً)",
                                  key="db_restore_c")
            if st.button("♻️ استعادة", disabled=not confirm, type="primary", key="db_restore_go"):
                ok, _ = run_safe(_restore_db, up.getvalue())
                if ok:
                    st.session_state.clear()
                    flash("تمت الاستعادة. سجّل الدخول مرة أخرى.", "success")
                    st.rerun()


def integrity_tab():
    st.caption("مراجعة محاسبية ومخزنية شاملة (قراءة فقط): توازن القيود، الفواتير مقابل بنودها والضريبة، المسدد مقابل السندات، "
               "أرصدة المخزون مقابل الحركات، الأرقام المكررة والمراجع المفقودة.")
    if st.button("🩺 تشغيل الفحص", type="primary", key="integ_run"):
        ok, res = run_safe(F.integrity_check)
        if ok:
            st.session_state["_integ"] = res
    res = st.session_state.get("_integ")
    if not res:
        return
    errs = sum(1 for r in res if r[0] == "خطأ")
    warns = sum(1 for r in res if r[0] == "تنبيه")
    m = st.columns(3)
    m[0].metric("أخطاء", errs)
    m[1].metric("تنبيهات", warns)
    m[2].metric("فحوص سليمة", len(res) - errs - warns)
    icon = {"خطأ": "✖ خطأ", "تنبيه": "⚠ تنبيه", "سليم": "✔ سليم"}
    table(pd.DataFrame([(icon[a], b, c, d or "", e or "") for a, b, c, d, e in res],
                       columns=["النتيجة", "المجال", "الفحص", "العدد", "أمثلة"]), height=480)
    c1, c2 = st.columns(2)
    if c1.button("🛠️ إعادة احتساب أرصدة وتكلفة المخزون", key="integ_fix_inv"):
        if run_safe(F.repair_inventory)[0]:
            st.session_state["_integ"] = F.integrity_check()
            flash("تم إعادة احتساب المخزون (مع نسخة احتياطية قبلها).")
            st.rerun()
    if c2.button("🛠️ إعادة احتساب المسدد وحالة الفواتير", key="integ_fix_paid"):
        if run_safe(F.repair_invoice_paid)[0]:
            st.session_state["_integ"] = F.integrity_check()
            flash("تم إعادة احتساب المسدد وحالة الفواتير (مع نسخة احتياطية قبلها).")
            st.rerun()


def _copy_db(path):
    src, dst = sqlite3.connect(F.DB_FILE), sqlite3.connect(path)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()


# ============================================================
# 📚 الدفتر العام (الترحيل الآلي) والقوائم المالية
# ============================================================
def page_ledger():
    header("📚 الدفتر العام والقوائم المالية")
    ok, _ = run_safe(F.sync_ledger)
    ok, stt = run_safe(F.ledger_status)
    if not ok:
        return
    m = st.columns(4)
    m[0].metric("قيود الترحيل الآلي", f"{stt['auto_entries']:,}")
    m[1].metric("قيود يومية (يدوي + خصم لاحق)", f"{stt['manual_entries']:,}")
    m[2].metric("إجمالي المدين = الدائن", money(stt["debit"]))
    m[3].metric("حالة الدفتر", "متوازن ✓" if stt["balanced"] else "غير متوازن ⚠")
    if stt["pending"]:
        st.info(f"{stt['pending']} مستند بانتظار الترحيل (يُرحَّل تلقائياً).")
    if stt["last_error"]:
        st.error(f"آخر خطأ ترحيل: {stt['last_error']}")
    if abs(stt["suspense_balance"]) > 0.005:
        st.warning(f"رصيد في الحساب المعلق {stt['suspense_code']}: {money(stt['suspense_balance'])} — سندات على خزينة/حساب "
                   "غير معرّف بالدليل. عرّف الحساب ثم أعد الترحيل.")
    with st.expander("📊 القيود الآلية حسب نوع المستند", expanded=False):
        table(pd.DataFrame(stt["by_source"], columns=["المصدر", "عدد القيود", "إجمالي المدين"]), money_cols=["إجمالي المدين"])
        st.caption(f"آخر إعادة ترحيل شاملة: {stt['rebuilt_at'] or '-'}. كل مستند يُرحَّل لحظة حفظه أو تعديله أو حذفه، "
                   "داخل نفس المعاملة.")
    if U()["is_admin"]:
        c1, c2 = st.columns(2)
        if c1.button("♻️ إعادة ترحيل كل المستندات", key="gl_rebuild", width="stretch"):
            if run_safe(F.rebuild_ledger)[0]:
                with F.db() as conn:
                    F.audit_event(conn, F.ACT_EVENT, "الدفتر العام", "إعادة ترحيل شاملة")
                flash("تمت إعادة ترحيل كل المستندات.")
                st.rerun()
        with c2.popover("⚙️ حسابات الترحيل الآلي", width="stretch"):
            with F.db() as conn:
                cur = F.gl_accounts(conn)
                accs = conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE COALESCE(is_group,0)=0 "
                                    "ORDER BY account_code").fetchall()
            names = dict(accs)
            codes = list(names)
            chosen = {}
            for key, (_d, label, _t, _p) in F.GL_ACCOUNT_DEFAULTS.items():
                chosen[key] = st.selectbox(label, codes, index=codes.index(cur[key]) if cur.get(key) in codes else None,
                                           format_func=lambda c: f"{c} - {names[c]}", key=f"glacc_{key}")
            if st.button("💾 حفظ وإعادة الترحيل", type="primary", key="glacc_save"):
                if run_safe(F.set_gl_accounts, {k: v for k, v in chosen.items() if v})[0]:
                    flash("تم حفظ الحسابات وإعادة ترحيل كل المستندات.")
                    st.rerun()
    keys = [k for k in ("bs", "gl_tb", "gl_pl", "gl_recon", "gl_journal") if can_report(k)]
    if not keys:
        return
    st.divider()
    key = st.radio("القائمة", keys, format_func=lambda k: F.REPORT_INDEX[k][0], horizontal=True, key="gl_rep_key",
                   label_visibility="collapsed")
    d1, d2 = date_range("gl", start=date(date.today().year, 1, 1))
    if d1 > d2:
        return
    ok, rep = run_safe(F.run_report, key, dict(d_from=d1, d_to=d2))
    if ok:
        render_report(rep, f"{key}_{d1}_{d2}")


# ============================================================
# 🧭 التخطيط المالي: التدفق النقدي، الموازنة، العمولات، الائتمان
# ============================================================
def page_planning():
    header("🧭 التخطيط المالي والائتمان")
    t1, t2, t3, t4 = st.tabs(["💧 توقع التدفق النقدي", "🎯 الموازنة التقديرية", "💼 عمولات المناديب", "🛡️ حدود الائتمان"])
    with t1:
        c1, c2, c3 = st.columns(3)
        to = c1.date_input("حتى تاريخ", date.today() + timedelta(days=90), key="cf_to", format="YYYY-MM-DD")
        gran = c2.selectbox("التجميع", ["شهري", "أسبوعي"], key="cf_gran")
        view = c3.selectbox("العرض", ["التوقع المجمع", "الفواتير المفتوحة بالاستحقاق"], key="cf_view")
        key = "cf_forecast" if view == "التوقع المجمع" else "due_inv"
        ok, rep = run_safe(F.run_report, key, dict(d_to=to.isoformat(), group="month" if gran == "شهري" else "day"))
        if ok:
            if key == "cf_forecast" and rep.rows:
                low = min(r[5] for r in rep.rows)
                m = st.columns(3)
                m[0].metric("تحصيلات متوقعة", money(sum(r[1] for r in rep.rows)))
                m[1].metric("مدفوعات متوقعة", money(sum(r[2] + r[3] for r in rep.rows)))
                m[2].metric("أدنى رصيد نقدي متوقع", money(low), delta="عجز" if low < 0 else None, delta_color="inverse")
            render_report(rep, f"{key}_{to}_{gran}", key=key)
    with t2:
        budget_editor()
        st.divider()
        d1, d2 = date_range("bva", start=date(date.today().year, 1, 1))
        if d1 <= d2:
            ok, rep = run_safe(F.run_report, "bva", dict(d_from=d1, d_to=d2))
            if ok:
                render_report(rep, f"bva_{d1}_{d2}")
    with t3:
        commissions_tab()
    with t4:
        if U()["is_admin"]:
            with st.form("credit_defaults"):
                c1, c2 = st.columns(2)
                pols = list(F.CREDIT_POLICIES)
                pol = c1.selectbox("السياسة الافتراضية عند التجاوز", pols, index=pols.index(F.default_policy()),
                                   format_func=F.CREDIT_POLICIES.get)
                days = c2.number_input("مدة السداد الافتراضية (يوم)", 0, 3650, F.default_credit_days())
                if st.form_submit_button("💾 حفظ الإعدادات", type="primary"):
                    if run_safe(F.set_credit_defaults, pol, int(days))[0]:
                        flash("تم حفظ إعدادات الائتمان.")
                        st.rerun()
        st.caption("الحد الائتماني ومدة السداد والسياسة لكل عميل من صفحة «العملاء». سياسة «المنع» تطبَّق في المحرك على أوامر "
                   "البيع وأذون التسليم في كل الواجهات، ويتجاوزها مدير النظام فقط.")
        if can_fin("customer"):
            ok, rep = run_safe(F.run_report, "credit", dict(d_from=F.today(), d_to=F.today()))
            if ok:
                render_report(rep, f"credit_{F.today()}", key="credit")


def budget_editor():
    st.markdown("#### ✏️ إدخال الموازنة الشهرية")
    choices = dict(F.budget_line_choices())
    c1, c2 = st.columns([1, 3])
    year = c1.number_input("السنة", 2000, 2100, date.today().year, key="bud_year")
    mat = F.budget_year_matrix(int(year))
    keys = list(dict.fromkeys([F.LINE_SALES, F.LINE_EXPENSES] + list(mat)))
    extra = c2.multiselect("إضافة بنود حسابات (إيرادات/مصروفات)", [k for k in choices if k not in keys],
                           format_func=choices.get, key="bud_extra")
    keys += extra
    months = [f"{m}" for m in F.AR_MONTHS]
    df = pd.DataFrame([[choices.get(k, k)] + [mat.get(k, [None] * 12)[i] for i in range(12)] for k in keys],
                      columns=["البند"] + months)
    edited = st.data_editor(df, key=f"bud_ed_{year}", hide_index=True, width="stretch", disabled=["البند"],
                            column_config={m: st.column_config.NumberColumn(m, min_value=0.0, format="%.2f") for m in months})
    if st.button("💾 حفظ الموازنة", type="primary", key="bud_save", disabled=not (U()["is_admin"] or can("planning"))):
        def save_all():
            for k, (_, row) in zip(keys, edited.iterrows()):
                vals = [None if pd.isna(row[m]) else float(row[m]) for m in months]
                old = mat.get(k, [None] * 12)
                for i, (v, o) in enumerate(zip(vals, old)):
                    per = f"{int(year):04d}-{i + 1:02d}"
                    if v is None and o is not None:
                        F.delete_budget(per, k)
                    elif v is not None and v != o:
                        F.save_budget(per, k, v, user=U()["username"])
        if run_safe(save_all)[0]:
            flash("تم حفظ الموازنة.")
            st.rerun()


# ============================================================
# 🛡️ سجل التدقيق والرقابة الداخلية
# ============================================================
def page_audit():
    header("🛡️ سجل التدقيق والرقابة الداخلية")
    ok, chain = run_safe(F.verify_audit_chain)
    if ok:
        good, n, bad = chain
        (st.success if good else st.error)(
            f"سلسلة البصمات سليمة ✓ — {n:,} سجل. السجل غير قابل للتعديل أو الحذف." if good else
            "⚠ تنبيه: تلاعب محتمل في السجل" + (f" عند السجل رقم {bad}" if bad else " (حماية السجل أُزيلت)") +
            ". راجع نسخ قاعدة البيانات الاحتياطية.")
    users, types = F.audit_filter_options()
    c1, c2 = st.columns(2)
    d1 = c1.date_input("من تاريخ", date.today() - timedelta(days=30), key="au_from", format="YYYY-MM-DD")
    d2 = c2.date_input("إلى تاريخ", date.today(), key="au_to", format="YYYY-MM-DD")
    c1, c2, c3, c4 = st.columns(4)
    user = c1.selectbox("المستخدم", [""] + users, format_func=lambda x: x or "الكل", key="au_user")
    dtype = c2.selectbox("نوع المستند", [""] + types, format_func=lambda x: x or "الكل", key="au_type")
    act = c3.selectbox("العملية", ["", F.ACT_UPDATE, F.ACT_DELETE, F.ACT_EVENT], format_func=lambda x: x or "الكل", key="au_act")
    q = c4.text_input("بحث (رقم المستند / قيمة)", key="au_q")
    ok, rep = run_safe(F.audit_report, d_from=d1.isoformat(), d_to=d2.isoformat(), user=user, doc_type=dtype, action=act, q=q)
    if not ok:
        return
    st.caption(f"عدد السجلات: {len(rep.rows):,} (الأحدث أولاً، حتى 2000 سجل).")
    render_report(rep, f"audit_{d1}_{d2}", height=520)


# ============================================================
# 💰 الخزينة والبنوك والسداد (مع «سداد هذه الفاتورة» الجاهز)
# ============================================================
def _voucher_parties(category):
    with F.db() as conn:
        if category == F.CAT_PARTNER:
            return [((c, n), f"{c} - {n}") for c, n in
                    conn.execute("SELECT partner_code, name FROM business_partners ORDER BY name")]
        if category == F.CAT_EXPENSE:
            pairs = [((c, n), f"{c} - {n}") for c, n in conn.execute(
                "SELECT account_code, account_name FROM chart_of_accounts WHERE account_type=? AND COALESCE(is_group,0)=0 "
                "ORDER BY account_code", (F.TYPE_EXP,)) if not c.startswith(F.NON_CASH_EXPENSE_PREFIXES)]
            return pairs or [((None, n), n) for n in F.DEFAULT_EXPENSES]
        if category == F.CAT_EMPLOYEE:
            return [((c, n), f"{c} - {n}") for c, n in F.employee_accounts(conn)]
        pairs = [((None, n), f"{c} - {n}") for c, n in conn.execute("SELECT asset_code, asset_name FROM fixed_assets")]
        return pairs + [((None, "أصول ثابتة - أخرى"), "أصول ثابتة - أخرى")]


def _open_invoices(kind, code):
    c = F.FLOW[kind]
    with F.db() as conn:
        rows = conn.execute(f"SELECT {c['iid']}, COALESCE(invoice_no, {c['iid']}), {c['idate']}, total_amount, paid_amount "
                            f"FROM {c['inv']} WHERE {c['pcode']}=? AND status='مفتوحة' ORDER BY {c['iid']}", (code,)).fetchall()
        out = {}   # {(النوع، المعرّف): (العنوان، المتبقي بعد المسدد والخصومات اللاحقة)}
        for i, no, d, t, p in rows:
            rem = round(t - p - F.invoice_discount_total(conn, kind, i), 2)
            out[(kind, i)] = (f"{'مبيعات' if kind == 'sales' else 'مشتريات'} {no} ({d}) - متبقي: {rem:,.2f}", max(rem, 0.0))
        return out


def _apply_cash_prefill():
    """تعبئة نموذج السند من «سداد هذه الفاتورة» قبل رسم العناصر (مرة واحدة): النوع، الطرف، الفاتورة، المبلغ، البيان."""
    pre = st.session_state.pop("cash_prefill", None)
    if not pre:
        return None
    st.session_state["cv_type"] = pre["t_type"]
    st.session_state["cv_cat"] = F.CAT_PARTNER
    st.session_state["cv_party"] = tuple(pre["party"])
    st.session_state["cv_inv"] = tuple(pre["alloc"])
    st.session_state["cv_amount"] = float(pre["remaining"])
    st.session_state["cv_desc"] = pre["description"]
    st.session_state.pop("cv_date", None)   # التاريخ الافتراضي = اليوم (بدون تحذير تكرار قيمة العنصر)
    st.session_state["cash_pay_info"] = dict(no=pre["no"], party=tuple(pre["party"]), alloc=tuple(pre["alloc"]),
                                             t_type=pre["t_type"])
    return pre


def _match_party(parties, want):
    """اختيار الطرف تلقائياً بالكود حتى لو تغيّر اسمه بعد إصدار الفاتورة."""
    if want in parties:
        return want
    code = want[0] if isinstance(want, (tuple, list)) and want else None
    return next((k for k in parties if code and k[0] == code), None)


def _on_cv_inv_change():
    """اختيار فاتورة من القائمة يعبّئ المبلغ المتبقي والبيان تلقائياً."""
    k = st.session_state.get("cv_inv")
    if not k:
        return
    try:
        pre = F.invoice_payment_prefill(*k)
    except F.UserError:
        return
    st.session_state["cv_amount"] = float(pre["remaining"])
    st.session_state["cv_desc"] = pre["description"]


def page_cash():
    header("💰 الخزينة والبنوك والسداد")
    back = st.session_state.get("cash_return")
    if back:
        kind, inv_id = back
        st.button(f"↩️ العودة لتفاصيل الفاتورة {F.invoice_number(kind, inv_id)}", key="cash_back",
                  on_click=_cash_go_back)
    voucher_form()
    cash_registers()


def _cash_go_back():
    kind, inv_id = st.session_state.pop("cash_return")
    go_to(kind, **{f"flow_view_{kind}": "inv", f"inv_inline_{kind}": inv_id})


def voucher_form(in_workspace=False):
    """نموذج سند القبض/الصرف (مشترك بين شاشة الخزينة وتبويب «السداد» في مساحة العمل).
    بيانات «سداد هذه الفاتورة» المحفوظة في الحالة تُطبَّق قبل رسم العناصر: النوع، الطرف، الفاتورة، المبلغ، البيان."""
    pre = _apply_cash_prefill()
    if pre and pre["remaining"] <= 0.005:
        st.info(f"✅ الفاتورة {pre['no']} مسددة بالكامل، لا يوجد مبلغ متبقٍ.")
    types = [t for t, k in (("قبض", "cash.receipt"), ("صرف", "cash.payment")) if can(k)]
    if types:
        with F.db() as conn:
            banks = F.cash_account_names(conn)
        st.markdown("#### ✍️ سند قبض / صرف")
        c1, c2, c3 = st.columns(3)
        if st.session_state.get("cv_type") not in types:
            st.session_state["cv_type"] = types[0]
        t_type = c1.selectbox("نوع السند", types, key="cv_type", format_func=lambda t: f"{t} ({'وارد' if t == 'قبض' else 'منصرف'})")
        bank = c2.selectbox("الخزينة / البنك", banks, key="cv_bank")
        t_date = c3.date_input("التاريخ", date.today(), key="cv_date", format="YYYY-MM-DD", max_value=date.today())
        c1, c2 = st.columns(2)
        cats = [F.CAT_PARTNER, F.CAT_EXPENSE, F.CAT_ASSET, F.CAT_EMPLOYEE]
        cat = c1.selectbox("التصنيف", cats, key="cv_cat")
        parties = dict(_voucher_parties(cat))
        if st.session_state.get("cv_party") is not None:
            m = _match_party(parties, st.session_state["cv_party"])
            if m is None:
                st.session_state.pop("cv_party", None)
            elif m != st.session_state["cv_party"]:
                st.session_state["cv_party"] = m
        party = c2.selectbox("الطرف / الحساب", list(parties), format_func=parties.get, index=None, key="cv_party",
                             placeholder="ابحث بالاسم أو الكود")
        alloc, invs = None, {}
        if cat == F.CAT_PARTNER and party and party[0]:
            invs = _open_invoices("sales" if t_type == "قبض" else "purchase", party[0])
            cur_inv = st.session_state.get("cv_inv")
            if cur_inv is not None and tuple(cur_inv) not in invs:
                st.session_state.pop("cv_inv", None)
            alloc = st.selectbox("الفاتورة للسداد (اختياري)", [None] + list(invs), key="cv_inv",
                                 on_change=_on_cv_inv_change,
                                 format_func=lambda k: "بدون فاتورة محددة (دفعة على الحساب)" if k is None else invs[k][0])
        if cat == F.CAT_PARTNER and party and party[0]:   # المندوب يُورَّث تلقائياً (لا يُختار يدوياً)
            with F.db() as conn:
                rep, src = F.voucher_rep_info(conn, party[0], t_date.isoformat(), *(tuple(alloc) if alloc else (None, None)))
            st.caption(f"👤 المندوب: **{rep}** ({src} — تلقائي)" if rep else "👤 المندوب: بدون مندوب (الطرف بلا مندوب مسجل)")
        info = st.session_state.get("cash_pay_info")
        if alloc and info and tuple(info["alloc"]) == tuple(alloc):
            st.success(f"💳 سند {info['t_type']} جاهز لسداد الفاتورة {info['no']} — {parties.get(party, info['party'][1])}. "
                       "كل البيانات معبأة تلقائياً: راجع الخزينة/البنك ثم اضغط «حفظ السند».")
        c1, c2 = st.columns([1, 2])
        amount = c1.number_input("المبلغ", min_value=0.0, step=100.0, format="%.2f", key="cv_amount")
        desc = c2.text_input("البيان", key="cv_desc")
        if alloc and tuple(alloc) in invs:
            rem = invs[tuple(alloc)][1]
            after = round(rem - (amount or 0), 2)
            with st.container(horizontal=True, key="cv_pay_summary"):
                st.metric("المتبقي على الفاتورة", money(rem), border=True)
                st.metric("مبلغ هذا السند", money(amount), border=True)
                st.metric("المتبقي بعد السداد", money(max(after, 0)),
                          "سداد كامل" if abs(after) < 0.005 else ("يتجاوز المتبقي!" if after < 0 else "سداد جزئي"),
                          delta_color="off", border=True)
        if st.button("💾 حفظ السند", type="primary", key="cv_save"):
            if F.period_is_closed(t_date.isoformat()) and not U()["is_admin"]:
                st.error("هذا التاريخ مغلق بالإغلاق اليومي. التسجيل فيه للمدير فقط.")
            else:
                ok, vid = run_safe(F.save_voucher, t_type, t_date.isoformat(), bank, cat, party, amount, alloc, desc)
                if ok:
                    for k in ("cv_amount", "cv_desc", "cv_inv", "cash_pay_info"):
                        st.session_state.pop(k, None)
                    flash(f"تم حفظ السند رقم {vid}" + (" وتسوية الفاتورة." if alloc else ".") +
                          (" تبويبات التقرير وكشف الحساب والفاتورة ما زالت مفتوحة أعلاه." if in_workspace else ""))
                    if in_workspace and "invoice" in (st.session_state.get("ws_open") or []):
                        st.session_state["ws_active"] = "invoice"   # عرض الفاتورة بعد السداد (حالتها المحدّثة)
                    st.rerun()


def cash_registers():
    if can("cash.register") or not any(k in U()["perms"] for k, _ in F.SUB_PERMISSIONS["cash"]):
        st.divider()
        with F.db() as conn:
            bal = F.cash_balances(conn)
            rows = conn.execute("SELECT trans_id, trans_date, trans_type, account_bank, category, party_name, amount, "
                                "COALESCE(description,'') FROM cash_transactions ORDER BY trans_date DESC, trans_id DESC "
                                "LIMIT 300").fetchall()
        st.markdown("#### 🏦 أرصدة الخزينة والبنوك")
        items = [(k, v) for k, v in bal.items() if abs(v) > 0.005] or [("لا توجد أرصدة", 0.0)]
        for i in range(0, len(items), 4):
            cols = st.columns(4)
            for col, (k, v) in zip(cols, items[i:i + 4]):
                col.metric(k, money(v))
        st.markdown("#### 📒 آخر السندات")
        table(pd.DataFrame(rows, columns=["رقم السند", "التاريخ", "النوع", "الحساب", "التصنيف", "الطرف", "المبلغ", "البيان"]),
              money_cols=["المبلغ"], height=420)


def go_pay_invoice(kind, inv_id):
    """زر «سداد هذه الفاتورة» خارج مساحة العمل: ينقل لشاشة الخزينة بكل البيانات جاهزة + زر رجوع للفاتورة."""
    pre = F.invoice_payment_prefill(kind, inv_id)
    back = (kind, inv_id) if st.session_state.get("nav") in ("sales", "purchase") else None
    go_to("cash", cash_prefill=pre, cash_pay=(kind, inv_id), cash_return=back)


# ============================================================
# الدورة المستندية المتصلة: أمر ← إذن ← فاتورة (للمراجعة والاعتماد)
# ============================================================
def _items_catalog():
    with F.db() as conn:
        return {c: (n, u or "", sp or 0.0, cp or 0.0, q or 0.0) for c, n, u, sp, cp, q in conn.execute(
            "SELECT item_code, item_name, unit, COALESCE(selling_price,0), COALESCE(cost_price,0), quantity "
            "FROM inventory ORDER BY item_name")}


def orders_section(kind):
    sales = kind == "sales"
    who = "العميل" if sales else "المورد"
    noun = "إذن تسليم" if sales else "إذن استلام"
    if can(f"{kind}.orders") and not can(f"{kind}.orders_create"):
        st.info("👁 عرض فقط: إنشاء الأوامر غير متاح لصلاحياتك.")
    if can(f"{kind}.orders_create"):
        cat = _items_catalog()
        parts = dict(partners("customer" if sales else "supplier"))
        st.markdown(f"#### 📝 {'أمر بيع' if sales else 'أمر شراء'} جديد")
        c1, c2, c3 = st.columns(3)
        party = c1.selectbox(who, list(parts), index=None, format_func=lambda c: f"{c} - {parts[c]}", key=f"ord_party_{kind}",
                             placeholder="اختر من القائمة")
        o_date = c2.date_input("تاريخ الأمر", date.today(), key=f"ord_date_{kind}", format="YYYY-MM-DD")
        exp = c3.date_input("موعد التنفيذ المتوقع", None, key=f"ord_exp_{kind}", format="YYYY-MM-DD")
        c1, c2 = st.columns(2)
        reps = [""] + F.rep_names()
        rep = c1.selectbox("المندوب" if sales else "المسؤول", reps, key=f"ord_rep_{kind}", format_func=lambda x: x or "—")
        notes = c2.text_input("ملاحظات", key=f"ord_notes_{kind}")
        cart_key = f"ord_cart_{kind}"
        cart = st.session_state.setdefault(cart_key, [])
        problems = []
        st.markdown("**➕ إضافة صنف**")
        c1, c2 = st.columns([3, 1])
        item = c1.selectbox("الصنف", list(cat), index=None, key=f"ord_item_{kind}", placeholder="ابحث بالاسم أو الكود",
                            format_func=lambda c: f"{c} - {cat[c][0]} (المتاح {cat[c][4]:g})")
        qty = c2.number_input("الكمية", min_value=0.0, step=1.0, format="%.3f", key=f"ord_qty_{kind}")
        c1, c2, c3 = st.columns([1, 1, 1])
        default_price = (cat[item][2] if sales else cat[item][3]) if item else 0.0
        price = c1.number_input("السعر" if sales else "التكلفة", min_value=0.0, value=float(default_price), format="%.2f",
                                key=f"ord_price_{kind}_{item}")
        disc = c2.text_input("الخصم الخاص (مبلغ أو 10%)", key=f"ord_disc_{kind}") if sales else ""
        if c3.button("➕ إضافة للأمر", key=f"ord_add_{kind}", width="stretch", disabled=not item or qty <= 0):
            try:
                d = F.parse_discount(disc, round(qty * price, 2)) if sales else 0.0
                cart.append(F.make_line(item, cat[item][0], cat[item][1], qty, price, d))
                st.session_state.pop(f"ord_item_{kind}", None)
                st.session_state.pop(f"ord_disc_{kind}", None)
                st.rerun()
            except F.UserError as e:
                st.error(str(e))
        lines = list(cart)
        if lines:
            table(pd.DataFrame([(i + 1, L["code"], L["name"], L["qty"], L["price"], L["discount"], L["subtotal"], L["vat"],
                                 L["total"]) for i, L in enumerate(lines)],
                               columns=["#", "الكود", "الصنف", "الكمية", "السعر", "الخصم", "قبل الضريبة", "الضريبة", "الإجمالي"]),
                  money_cols=["السعر", "الخصم", "قبل الضريبة", "الضريبة", "الإجمالي"], qty_cols=["الكمية"])
            c1, c2 = st.columns([3, 1])
            rm = c1.selectbox("حذف سطر", list(range(len(lines))), index=None, key=f"ord_rm_{kind}",
                              format_func=lambda i: f"{i + 1}) {lines[i]['name']} × {lines[i]['qty']:g}",
                              placeholder="اختر سطراً لحذفه", label_visibility="collapsed")
            if c2.button("🗑️ حذف", key=f"ord_rm_go_{kind}", disabled=rm is None, width="stretch"):
                cart.pop(rm)
                st.session_state.pop(f"ord_rm_{kind}", None)
                st.rerun()
        else:
            st.caption("لم تُضف أصناف بعد. السعر الافتراضي من بطاقة الصنف ويمكن تعديله.")
        for p in problems:
            st.error(p)
        if lines:
            sub, vat = sum(l["subtotal"] for l in lines), sum(l["vat"] for l in lines)
            st.caption(f"عدد الأصناف: {len(lines)}   |   قبل الضريبة: {money(sub)}   |   الضريبة: {money(vat)}   |   "
                       f"الإجمالي: {money(sub + vat)} ر.س")
            if sales and party:
                cst = F.credit_status(party, round(sub + vat, 2), "order")
                if cst["exceeded"]:
                    (st.error if cst["policy"] == F.POLICY_BLOCK else st.warning)("⚠ " + cst["message"].replace("\n", "  \n"))
        override = False
        if sales and U()["is_admin"]:
            override = st.checkbox("موافقة المدير على تجاوز الحد الائتماني (عند سياسة المنع)", key=f"ord_over_{kind}")
        if st.button(f"💾 حفظ {'أمر البيع' if sales else 'أمر الشراء'}", type="primary", key=f"ord_save_{kind}",
                     disabled=not (party and lines) or bool(problems)):
            ok, oid = run_safe(F.create_order, kind, o_date.isoformat(), (party, parts[party]), notes, lines,
                               exp.isoformat() if exp else None, rep, override)
            if ok:
                st.session_state[f"ord_cart_{kind}"] = []
                st.session_state[f"ord_last_{kind}"] = oid
                flash(f"تم حفظ الأمر رقم {oid}. المخزن لم يتأثر بعد.")
                st.rerun()
        last = st.session_state.get(f"ord_last_{kind}")
        if last and can(f"{kind}.{'delivery' if sales else 'receipt'}"):
            st.button(f"➡️ ترحيل الأمر {last} إلى {noun} (كل الأصناف والكميات تلقائياً)", type="primary",
                      key=f"ord_go_last_{kind}", on_click=go_to, args=(kind,),
                      kwargs={f"flow_view_{kind}": "ful", f"ful_order_{kind}": last, f"ord_last_{kind}": None})
        st.divider()
    c = F.FLOW[kind]
    with F.db() as conn:
        rows = conn.execute(f"SELECT o.order_id, o.order_date, o.{c['pname']}, o.status, o.total_amount, "
                            f"COALESCE(SUM(l.quantity),0), COALESCE(SUM(l.{c['done']}),0) FROM {c['order']} o "
                            f"LEFT JOIN {c['olines']} l ON l.order_id=o.order_id GROUP BY o.order_id "
                            f"ORDER BY o.order_id DESC LIMIT 300").fetchall()
    st.markdown(f"#### 📋 {'أوامر البيع' if sales else 'أوامر الشراء'}")
    df = pd.DataFrame([(r[0], r[1], r[2], r[3], r[4], (r[6] / r[5] * 100) if r[5] else 0.0) for r in rows],
                      columns=["رقم الأمر", "التاريخ", who, "الحالة", "الإجمالي", "% المنفّذ"])
    sel = table(df, money_cols=["الإجمالي"], pct_cols=["% المنفّذ"], height=320, select_key=f"ord_tbl_{kind}")
    if sel:
        oid, status = rows[sel[0]][0], rows[sel[0]][3]
        if status not in F.OPEN_BLOCKED and can(f"{kind}.{'delivery' if sales else 'receipt'}"):
            st.button(f"➡️ ترحيل الأمر {oid} إلى {noun}", type="primary", key=f"ord_go_{kind}", on_click=go_to,
                      args=(kind,), kwargs={f"flow_view_{kind}": "ful", f"ful_order_{kind}": oid, f"ord_tbl_{kind}": None})
        else:
            st.caption(f"الأمر {oid} حالته «{status}» ولا يمكن تنفيذه.")


def fulfil_section(kind):
    sales = kind == "sales"
    noun = "إذن التسليم" if sales else "إذن الاستلام"
    c = F.FLOW[kind]
    ph = ",".join("?" * len(F.OPEN_BLOCKED))
    with F.db() as conn:
        orders = conn.execute(f"SELECT order_id, order_date, {c['pname']}, status FROM {c['order']} WHERE status NOT IN ({ph}) "
                              f"ORDER BY order_id DESC", F.OPEN_BLOCKED).fetchall()
    opts = {o[0]: f"أمر {o[0]} — {o[2]} ({o[1]}) — {o[3]}" for o in orders}
    key = f"ful_order_{kind}"
    if st.session_state.get(key) not in opts:
        st.session_state.pop(key, None)
    st.markdown(f"#### {'🚚' if sales else '📥'} {noun} من أمر")
    oid = st.selectbox("الأمر", list(opts), format_func=opts.get, index=None, key=key,
                       placeholder="اختر أمراً مفتوحاً (أو استخدم «ترحيل إلى إذن» من شاشة الأوامر)")
    if oid:
        with F.db() as conn:
            rows = conn.execute(f"SELECT l.line_id, l.item_code, l.item_name, l.quantity, l.{c['done']}, l.{c['oprice']}, "
                                f"COALESCE(i.quantity,0) FROM {c['olines']} l LEFT JOIN inventory i ON i.item_code=l.item_code "
                                f"WHERE l.order_id=? ORDER BY l.line_id", (oid,)).fetchall()
        st.caption("الكميات معبأة تلقائياً بالمتبقي من الأمر" + (" (بحد أقصى المتاح بالمخزن)" if sales else
                   " (السعر مقفول على سعر أمر الشراء)") + ". عدّل «الكمية الآن» للتنفيذ الجزئي.")
        inputs = []
        for lid, code, name, q, done, price, stock in rows:
            rem = max(q - done, 0.0)
            now = min(rem, max(stock, 0.0)) if sales else rem
            with st.container(border=True):
                info = (f"**{code} - {name}**  \nالمطلوب {q:g}  •  المنفّذ {done:g}  •  المتبقي {rem:g}  •  "
                        + (f"المتاح بالمخزن {stock:g}" if sales else f"السعر {money(price)} 🔒"))
                if sales:
                    c1, c2 = st.columns([3, 1])
                    c1.markdown(info)
                    qv = c2.number_input("الكمية الآن", 0.0, float(rem), float(now), format="%.3f",
                                         key=f"ful_q_{kind}_{oid}_{lid}", disabled=rem <= 0)
                    inputs.append((lid, qv, None, None))
                else:
                    st.markdown(info)
                    c1, c2, c3 = st.columns(3)
                    qv = c1.number_input("الكمية الآن", 0.0, float(rem), float(now), format="%.3f",
                                         key=f"ful_q_{kind}_{oid}_{lid}", disabled=rem <= 0)
                    pd_ = c2.date_input("تاريخ الإنتاج", None, key=f"ful_p_{kind}_{oid}_{lid}", format="YYYY-MM-DD")
                    ed_ = c3.date_input("تاريخ الانتهاء", None, key=f"ful_e_{kind}_{oid}_{lid}", format="YYYY-MM-DD")
                    inputs.append((lid, qv, pd_.isoformat() if pd_ else None, ed_.isoformat() if ed_ else None))
        c1, c2 = st.columns(2)
        f_date = c1.date_input("تاريخ " + noun, date.today(), key=f"ful_date_{kind}", format="YYYY-MM-DD")
        notes = c2.text_input("ملاحظات", key=f"ful_notes_{kind}")
        ref = st.text_input("رقم مستند المورد", key=f"ful_ref_{kind}") if not sales else ""
        override = sales and U()["is_admin"] and st.checkbox("موافقة المدير على تجاوز الحد الائتماني (عند سياسة المنع)",
                                                              key=f"ful_over_{kind}")
        create_ok = can(f"{kind}.{'delivery' if sales else 'receipt'}_create")
        ups = None
        if not sales and create_ok:
            ups = st.file_uploader("📎 فاتورة المورد الورقية (صورة / PDF) — تُحفظ مع إذن الاستلام", accept_multiple_files=True,
                                   key=f"ful_files_{kind}",
                                   type=["pdf", "png", "jpg", "jpeg", "gif", "bmp", "webp", "tif", "tiff", "heic"])
        if not create_ok:
            st.info("👁 عرض فقط: التنفيذ والاعتماد غير متاحين لصلاحياتك.")
        elif st.button(f"💾 حفظ {noun}", type="primary", key=f"ful_save_{kind}"):
            entries = [(lid, float(qv or 0), None, pdv, edv) for lid, qv, pdv, edv in inputs]
            if F.period_is_closed(f_date.isoformat()) and not U()["is_admin"]:
                st.error("هذا التاريخ مغلق بالإغلاق اليومي.")
            else:
                ok, fid = run_safe(F.fulfil_order, kind, oid, f_date.isoformat(), notes, entries, ref, bool(override))
                if ok:
                    n_att = 0
                    for up in ups or []:
                        ok_a, _ = run_safe(F.add_receipt_attachment, fid, up.name, up.getvalue())
                        n_att += bool(ok_a)
                    st.session_state[f"ful_last_{kind}"] = fid
                    st.session_state.pop(key, None)
                    st.session_state.pop(f"ful_files_{kind}", None)
                    flash(f"تم حفظ {noun} رقم {fid} وتحديث المخزن." + (f" 📎 أُرفق {n_att} ملف لفاتورة المورد." if n_att else ""))
                    st.rerun()
    last = st.session_state.get(f"ful_last_{kind}")
    if last and can(f"{kind}.invoices"):
        st.button(f"🧾 إصدار فاتورة من {noun} رقم {last} (للمراجعة والاعتماد)", type="primary", key=f"ful_inv_last_{kind}",
                  on_click=go_to, args=(kind,), kwargs={f"flow_view_{kind}": "review", f"review_ful_{kind}": last,
                                                         f"ful_last_{kind}": None})
    st.divider()
    with F.db() as conn:
        regs = conn.execute(f"SELECT f.{c['fid']}, f.{c['fdate']}, f.order_id, f.{c['pname']}, COUNT(l.line_id), "
                            f"COALESCE(SUM({c['fval']}),0), f.invoice_id FROM {c['ful']} f LEFT JOIN {c['flines']} l "
                            f"ON l.{c['fid']}=f.{c['fid']} GROUP BY f.{c['fid']} ORDER BY f.{c['fid']} DESC LIMIT 300").fetchall()
    st.markdown(f"#### 📒 سجل {'أذون التسليم' if sales else 'أذون الاستلام'}")
    df = pd.DataFrame([(r[0], r[1], r[2], r[3], r[4], r[5], F.invoice_number(kind, r[6]) if r[6] else "غير مفوتر")
                       for r in regs], columns=["رقم الإذن", "التاريخ", "رقم الأمر", "الطرف", "البنود", "القيمة", "الفاتورة"])
    sel = table(df, money_cols=["القيمة"], height=320, select_key=f"ful_tbl_{kind}")
    if sel and not regs[sel[0]][6] and can(f"{kind}.invoices"):
        st.button(f"🧾 إصدار فاتورة من الإذن {regs[sel[0]][0]}", type="primary", key=f"ful_inv_{kind}", on_click=go_to,
                  args=(kind,), kwargs={f"flow_view_{kind}": "review", f"review_ful_{kind}": regs[sel[0]][0],
                                        f"ful_tbl_{kind}": None})
    if sel and not sales:   # مرفقات فاتورة المورد على الإذن المحدد
        with st.container(border=True):
            st.caption(f"إذن الاستلام رقم {regs[sel[0]][0]} — {regs[sel[0]][3]}")
            attachments_box("receipt", regs[sel[0]][0])


def invoice_review_section(kind):
    """مسودة الفاتورة من الإذن: البنود والضريبة والإجمالي جاهزة للمراجعة، ثم «اعتماد وإصدار»."""
    sales = kind == "sales"
    fid = st.session_state.get(f"review_ful_{kind}")
    c = F.FLOW[kind]
    with F.db() as conn:
        h = conn.execute(f"SELECT {c['pcode']}, {c['pname']}, invoice_id, {c['fdate']} FROM {c['ful']} WHERE {c['fid']}=?",
                         (fid,)).fetchone() if fid else None
        if h:
            raw = (conn.execute("SELECT item_code, item_name, unit, quantity, unit_price, COALESCE(discount_amount,0) "
                                "FROM sales_delivery_lines WHERE delivery_id=? ORDER BY line_id", (fid,)).fetchall() if sales else
                   conn.execute("SELECT item_code, item_name, unit, quantity, unit_cost, 0 FROM purchase_receipt_lines "
                                "WHERE receipt_id=? ORDER BY line_id", (fid,)).fetchall())
    if not h:
        st.info("اختر إذناً من «الأذون» ثم «إصدار فاتورة».")
        return
    if h[2]:
        st.success(f"الإذن {fid} مفوتر بالفعل (فاتورة {F.invoice_number(kind, h[2])}).")
        return
    lines = [F.make_line(*r) for r in raw]
    st.markdown(f"#### 🧾 مراجعة {'الفاتورة الضريبية' if sales else 'فاتورة المورد'} قبل الاعتماد")
    st.caption(f"من {'إذن التسليم' if sales else 'إذن الاستلام'} رقم {fid} ({h[3]})   |   {h[0]} - {h[1]}")
    table(pd.DataFrame([(L["code"], L["name"], L["qty"], L["price"], L["discount"], L["subtotal"], L["vat"], L["total"])
                        for L in lines], columns=["الكود", "الصنف", "الكمية", "السعر", "الخصم", "قبل الضريبة",
                                                  "الضريبة 15%", "الإجمالي"]),
          money_cols=["السعر", "الخصم", "قبل الضريبة", "الضريبة 15%", "الإجمالي"], qty_cols=["الكمية"])
    sub, vat = sum(L["subtotal"] for L in lines), sum(L["vat"] for L in lines)
    m = st.columns(3)
    m[0].metric("قبل الضريبة", money(sub))
    m[1].metric("الضريبة 15%", money(vat))
    m[2].metric("الإجمالي", money(sub + vat))
    if sales:
        cst = F.credit_status(h[0], 0.0, "invoice")
        if cst["exceeded"]:
            st.warning("⚠ " + cst["message"].replace("\n", "  \n"))
    c1, c2, c3 = st.columns(3)
    d = c1.date_input("تاريخ الفاتورة", date.today(), key=f"rv_date_{kind}", format="YYYY-MM-DD")
    notes = c2.text_input("ملاحظات", key=f"rv_notes_{kind}")
    ref = c3.text_input("رقم فاتورة المورد", key=f"rv_ref_{kind}") if not sales else ""
    if st.button("✅ اعتماد وإصدار الفاتورة", type="primary", key=f"rv_go_{kind}"):
        ok, res = run_safe(F.invoice_unbilled_selection, kind, [fid], d.isoformat(), notes, ref, allow_closed=U()["is_admin"])
        if ok:
            st.session_state.pop(f"review_ful_{kind}", None)
            st.session_state[f"inv_open_{kind}"] = res[0]
            flash(f"تم إصدار الفاتورة رقم {F.invoice_number(kind, res[0])} بإجمالي {money(res[1])} ر.س")
            st.session_state[f"_flow_goto_{kind}"] = "inv"   # يُطبق قبل رسم محدد الأقسام في التحديث التالي
            st.rerun()


# ============================================================
# 💼 عمولات المناديب التفاعلية (نسبة موحدة أو تعديل مباشر لكل مندوب)
# ============================================================
def commissions_tab():
    c1, c2 = st.columns(2)
    d1 = c1.date_input("من تاريخ", date.today().replace(day=1), key="com_from", format="YYYY-MM-DD")
    d2 = c2.date_input("إلى تاريخ", date.today(), key="com_to", format="YYYY-MM-DD")
    if d1 > d2:
        st.warning("تاريخ البداية بعد تاريخ النهاية.")
        return
    c1, c2, c3 = st.columns(3)
    bases = list(F.COMMISSION_BASES)
    base = c1.selectbox("أساس الاحتساب", bases, index=bases.index(F.commission_base()), format_func=F.COMMISSION_BASES.get,
                        key="com_base")
    use_uniform = c2.toggle("نسبة موحدة للكل", key="com_use_uni")
    uniform = c3.number_input("النسبة الموحدة %", 0.0, 100.0, 2.0, 0.25, key="com_uni", disabled=not use_uniform)
    uniform = float(uniform) if use_uniform else None
    rates = st.session_state.setdefault("com_rates", {})
    p = dict(d_from=d1.isoformat(), d_to=d2.isoformat())
    ok, res = run_safe(F.commission_rows, p, rates, uniform, base)
    if not ok:
        return
    rows, base = res
    rows = sorted(rows, key=lambda r: r[1])   # ترتيب ثابت بالاسم حتى تبقى التعديلات على نفس الصف
    df = pd.DataFrame([(r[1], r[2], r[3], r[4], r[5], r[6], r[7]) for r in rows],
                      columns=["المندوب", "نسبة العمولة %", "صافي المبيعات", "التحصيلات", "أساس الاحتساب", "العمولة المستحقة",
                               "ملاحظات"])
    ed_key = f"com_ed_{d1}_{d2}_{base}_{uniform}"

    def on_edit():
        changes = st.session_state[ed_key].get("edited_rows", {})
        for idx, ch in changes.items():
            if "نسبة العمولة %" in ch and int(idx) < len(rows):
                v = ch["نسبة العمولة %"]
                rates[rows[int(idx)][1]] = None if v is None else float(v)

    def clear_edits():
        st.session_state["com_rates"] = {}
        st.session_state.pop(ed_key, None)

    def save_rates():
        ok, n = run_safe(F.save_commission_rates, {r[1]: r[2] for r in rows})
        if ok:
            if U()["is_admin"]:
                F.set_commission_base(base)
            clear_edits()
            flash(f"تم حفظ نسب {n} مندوب في سجل المناديب.")

    m = st.columns(3)
    m[0].metric("إجمالي العمولات المستحقة", money(df["العمولة المستحقة"].sum()))
    m[1].metric("إجمالي أساس الاحتساب", money(df["أساس الاحتساب"].sum()))
    m[2].metric("عدد المناديب", len(df))
    st.caption("✏️ عدّل «نسبة العمولة %» مباشرة أمام أي مندوب؛ العمولة والإجماليات تُحتسب فوراً. "
               "التعديل المباشر يتقدم على النسبة الموحدة.")
    edit_df = df[["المندوب", "نسبة العمولة %", "العمولة المستحقة"]]
    st.data_editor(edit_df, key=ed_key, hide_index=True, width="stretch", on_change=on_edit,
                   disabled=["المندوب", "العمولة المستحقة"],
                   column_config={"المندوب": st.column_config.Column(width="medium", alignment="right"),
                                  "نسبة العمولة %": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=0.25,
                                                                                  format="%.2f", alignment="right"),
                                  "العمولة المستحقة": st.column_config.NumberColumn(format="%.2f", alignment="right")})
    with st.expander("📊 التفاصيل الكاملة (المبيعات، التحصيلات، أساس الاحتساب)", expanded=True):
        table(df, money_cols=["صافي المبيعات", "التحصيلات", "أساس الاحتساب", "العمولة المستحقة"],
              pct_cols=[], qty_cols=["نسبة العمولة %"])
    c1, c2 = st.columns(2)
    c1.button("↺ إلغاء التعديلات", key="com_reset", width="stretch", on_click=clear_edits)
    c2.button("💾 حفظ النسب في سجل المناديب", key="com_save", width="stretch", on_click=save_rates,
              disabled=not (U()["is_admin"] or can("reps")))
    rep = F.run_report("commissions", dict(p, rates=rates, uniform=uniform, base=base))
    c1, c2 = st.columns(2)
    with c1:
        download_on_demand("Excel", f"com_x_{d1}_{d2}_{base}_{uniform}_{sorted(rates.items())}", F.export_report_xlsx, ".xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "commissions.xlsx", rep)
    with c2:
        download_on_demand("PDF", f"com_p_{d1}_{d2}_{base}_{uniform}_{sorted(rates.items())}", F.build_report_pdf, ".pdf",
                           "application/pdf", "commissions.pdf", rep)


# ============================================================
# التنقل الرئيسي
# ============================================================
PAGES = [
    ("dashboard", "🏠 لوحة المؤشرات", page_dashboard),
    ("customers", "👥 العملاء", lambda: page_partners("customer")),
    ("suppliers", "🏭 الموردون", lambda: page_partners("supplier")),
    ("reps", "🧑‍💼 سجل المناديب", page_reps),
    ("import", "📥 مركز استيراد البيانات", page_import),
    ("inventory", "🏭 المخزون وكارت الصنف", page_inventory),
    ("sales", "🛒 المبيعات والإشعارات الدائنة", lambda: page_flow("sales")),
    ("purchase", "📦 المشتريات وخصم الموردين", lambda: page_flow("purchase")),
    ("cash", "💰 الخزينة والبنوك والسداد", page_cash),
    ("statement", "📜 كشف الحساب الموحد", page_statement),
    ("discounts", "🏷️ الخصومات اللاحقة والإشعارات", page_discounts),
    ("income", "📈 قائمة الدخل", page_income),
    ("reports", "📊 التقارير", page_reports),
    ("journal", "📓 القيود اليومية", page_journal),
    ("ledger", "📚 الدفتر العام والقوائم المالية", page_ledger),
    ("planning", "🧭 التخطيط المالي والائتمان", page_planning),
    ("coa", "🗂️ دليل الحسابات", page_coa),
    ("company", "⚙️ الإعدادات", page_settings),
    ("users", "👤 المستخدمون والصلاحيات", page_users),
    ("audit", "🛡️ سجل التدقيق والرقابة", page_audit),
]
NAV_GROUPS = [("الرئيسية", ["dashboard"]), ("البيانات والمخزون", ["import", "customers", "suppliers", "reps", "inventory"]),
              ("العمليات", ["sales", "purchase", "cash"]),
              ("الحسابات", ["statement", "discounts", "income", "reports", "journal", "ledger", "coa"]),
              ("التخطيط", ["planning"]),
              ("الإدارة", ["company", "users", "audit"])]


def sidebar_brand():
    uri, bg = logo_data_uri(300)
    img = f'<img src="{uri}" alt="ELHAGRY"/>' if uri else ""
    st.markdown(f'<div class="fz-side-brand" style="background:radial-gradient(circle at 50% 30%, #1b4592 0%, {bg} 72%)">'
                f'{img}<div class="ar">{F.BRAND_AR}</div><div class="en">{F.BRAND_EN.upper()}</div></div>',
                unsafe_allow_html=True)


def main():
    boot()
    if not check_session():
        login_page()
        return
    u = U()
    st.session_state["_tbl_seq"] = 0   # مفاتيح حاويات الجداول ثابتة بين التحديثات
    keep_state_alive()                 # الفلاتر والطرف وبيانات السند لا تضيع عند التنقل بين الشاشات
    F.set_current_user(u["username"], "web", u)   # اسم المستخدم في سجل التدقيق لكل عملية في هذه الجلسة
    allowed = {k: (t, fn) for k, t, fn in PAGES if can(k)}
    current = st.session_state.get("nav")
    if current not in allowed:
        current = next(iter(allowed), None)
        st.session_state["nav"] = current
    with st.sidebar:
        sidebar_brand()
        initial = html.escape((u["full_name"] or u["username"])[:1])
        st.markdown(f'<div class="fz-user-chip"><div class="av">{initial}</div><div><div class="n">'
                    f'{html.escape(u["full_name"])}</div><div class="r">{"مدير النظام" if u["is_admin"] else "مستخدم"} • '
                    f'{html.escape(u["username"])}</div></div></div>', unsafe_allow_html=True)
        if not allowed:
            st.warning("لا توجد صلاحيات لهذا المستخدم.")
        for group, keys in NAV_GROUPS:
            keys = [k for k in keys if k in allowed]
            if not keys:
                continue
            st.markdown(f'<div class="fz-nav-h">{group}</div>', unsafe_allow_html=True)
            for k in keys:
                st.button(allowed[k][0], key=f"nav_btn_{k}", width="stretch",
                          type="primary" if k == current else "secondary", on_click=go_to, args=(k,))
        st.divider()
        if st.button("🚪 تسجيل الخروج", width="stretch", key="logout_btn"):
            st.session_state.clear()
            st.rerun()
        st.caption(f"© {F.BRAND_FULL}")
    if current:
        allowed[current][1]()


main()
