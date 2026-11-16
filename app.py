# -*- coding: utf-8 -*-
"""Alassal Smart Merchant - Arabic Streamlit demo."""
from __future__ import annotations
from pathlib import Path
from datetime import date, timedelta
from io import BytesIO
import math
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = APP_DIR / "Alassal_Wholesale_Synthetic_Data.xlsx"
REQUIRED = [
    "Products", "Suppliers", "Branches_Warehouses", "Sales_Reps", "Customers",
    "Purchase_Orders", "Purchase_Order_Items", "Sales_Invoices", "Sales_Invoice_Items",
    "Payments_Collections", "Inventory_Batches", "Returns"
]

st.set_page_config(page_title="العسال | Smart Merchant", page_icon="🧠", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], [data-testid="stAppViewContainer"] {font-family:'Cairo',sans-serif;}
.block-container {padding-top:1.2rem; padding-bottom:2rem; direction:rtl;}
[data-testid="stSidebar"] {direction:rtl;}
[data-testid="stMetric"] {background:#111827;padding:14px 16px;border:1px solid #263244;border-radius:14px;}
[data-testid="stMetricLabel"] {justify-content:flex-end;}
[data-testid="stMetricValue"] {direction:rtl;}
div[data-testid="stDataFrame"] {direction:rtl;}
.hero {background:linear-gradient(120deg,#101827,#173b4b);padding:24px 28px;border-radius:18px;border:1px solid #2c5263;margin-bottom:18px;}
.hero h1 {font-size:30px;margin:0;color:#fff;font-weight:800;}
.hero p {color:#d4e6eb;margin:8px 0 0;font-size:15px;}
.section-title {font-size:21px;font-weight:800;margin:10px 0 14px;color:var(--text-color);}
.small-note {font-size:12px;color:#9ca3af;}
.recommend {border:1px solid #334155;border-right:5px solid #0ea5a4;border-radius:12px;padding:13px 16px;margin:8px 0;background:rgba(15,23,42,.35);}
.warn {border-right-color:#f59e0b;} .danger {border-right-color:#ef4444;}
</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def read_workbook(source):
    xls = pd.ExcelFile(source)
    missing = [s for s in REQUIRED if s not in xls.sheet_names]
    if missing: raise ValueError("Sheets missing: " + ", ".join(missing))
    tables = {s: pd.read_excel(source, sheet_name=s) for s in REQUIRED}
    for df in tables.values():
        df.columns = [str(c).strip() for c in df.columns]
    for s, cols in {
        "Products":["unit_cost_EGP","unit_price_EGP","shelf_life_days","reorder_point_units","reorder_qty_units"],
        "Inventory_Batches":["batch_qty_units","days_to_expiry"],
        "Sales_Invoice_Items":["quantity","unit_price_EGP","discount_pct","line_gross_EGP","line_net_EGP"],
        "Sales_Invoices":["total_amount_EGP"],
        "Payments_Collections":["amount_paid_EGP"],
        "Purchase_Order_Items":["quantity","unit_cost_EGP"],
        "Returns":["quantity_returned"]
    }.items():
        for c in cols:
            if c in tables[s]: tables[s][c] = pd.to_numeric(tables[s][c], errors="coerce").fillna(0)
    for s in ["Purchase_Orders","Sales_Invoices","Payments_Collections","Inventory_Batches","Returns","Sales_Reps"]:
        for c in tables[s].columns:
            if any(k in c.lower() for k in ["date","hire"]): tables[s][c] = pd.to_datetime(tables[s][c], errors="coerce")
    return tables

@st.cache_data(show_spinner=False)
def derive(tables, as_of):
    p=tables["Products"].copy(); b=tables["Inventory_Batches"].copy(); si=tables["Sales_Invoice_Items"].copy()
    inv=tables["Sales_Invoices"].copy(); pay=tables["Payments_Collections"].copy()
    loc=tables["Branches_Warehouses"].copy(); cust=tables["Customers"].copy()
    # Batch-level expiry is recalculated from the actual selected date, not the stale source field.
    b["expiry_date"] = pd.to_datetime(b["expiry_date"], errors="coerce")
    b["days_left"]=(b["expiry_date"].dt.date - as_of).apply(lambda x: x.days if pd.notna(x) else None)
    b=b.merge(p[["product_id","product_name","category","unit","unit_cost_EGP","unit_price_EGP","reorder_point_units","reorder_qty_units","primary_supplier_id"]],on="product_id",how="left")
    b=b.merge(loc[["location_id","location_name"]],on="location_id",how="left")
    stock=b.groupby("product_id",as_index=False)["batch_qty_units"].sum().rename(columns={"batch_qty_units":"stock_qty"})
    sales=si.merge(inv[["invoice_id","order_date","customer_id","branch_id"]],on="invoice_id",how="left")
    sales["order_date"]=pd.to_datetime(sales["order_date"],errors="coerce")
    anchor=sales["order_date"].max()
    if pd.isna(anchor): anchor=pd.Timestamp(as_of)
    window_start=anchor.normalize()-pd.Timedelta(days=59)
    recent=sales[sales.order_date>=window_start]
    sold=recent.groupby("product_id",as_index=False).agg(sold_60d=("quantity","sum"),sales_value_60d=("line_net_EGP","sum"))
    last=sales.groupby("product_id",as_index=False).agg(last_sale_date=("order_date","max"))
    prod=p.merge(stock,on="product_id",how="left").merge(sold,on="product_id",how="left").merge(last,on="product_id",how="left")
    prod[["stock_qty","sold_60d","sales_value_60d"]]=prod[["stock_qty","sold_60d","sales_value_60d"]].fillna(0)
    prod["daily_sales_rate"]=prod["sold_60d"]/60
    prod["stock_cover_days"]=prod.apply(lambda r: (r.stock_qty/r.daily_sales_rate) if r.daily_sales_rate>0 else (999 if r.stock_qty>0 else 0),axis=1)
    prod["days_since_sale"]=(anchor.normalize()-prod["last_sale_date"]).dt.days
    prod["estimated_gross_profit"] = (prod["sales_value_60d"] - prod["sold_60d"]*prod["unit_cost_EGP"])
    prod["estimated_margin_pct"] = prod.apply(lambda r: (r.estimated_gross_profit/r.sales_value_60d*100) if r.sales_value_60d>0 else 0,axis=1)
    prod["stock_value_cost"] = prod["stock_qty"]*prod["unit_cost_EGP"]
    prod["reorder_flag"]=prod["stock_qty"]<prod["reorder_point_units"]
    # Receivables: aggregate payments by invoice before joining to invoice headers.
    paid=pay.groupby("invoice_id",as_index=False)["amount_paid_EGP"].sum().rename(columns={"amount_paid_EGP":"paid_EGP"})
    ar=inv.merge(paid,on="invoice_id",how="left").merge(cust[["customer_id","customer_name","customer_type"]],on="customer_id",how="left")
    ar["paid_EGP"]=ar["paid_EGP"].fillna(0)
    ar["balance_EGP"]=(ar["total_amount_EGP"]-ar["paid_EGP"]).clip(lower=0)
    ar["due_date"]=pd.to_datetime(ar["due_date"],errors="coerce")
    ar["days_overdue"]=(pd.Timestamp(as_of)-ar["due_date"]).dt.days
    ar["is_overdue"]=(ar.balance_EGP>0)&(ar.due_date.dt.date<as_of)
    # sales overview
    gross=si["line_gross_EGP"].sum(); net=si["line_net_EGP"].sum()
    return {"products":prod,"batches":b,"sales":sales,"recent_sales":recent,"ar":ar,"gross_sales":gross,"net_sales":net,
            "discount_value":gross-net,"anchor_date":anchor,"window_start":window_start}

def money(v): return f"{float(v):,.0f} ج.م"
def pct(v): return f"{float(v):.1f}%"
def max_safe_discount(price,cost,min_margin):
    if price<=0: return 0.0
    return max(0.0,min(50.0,(1-(cost/price)-min_margin/100)*100))

def recommendation_engine(d, tables, min_margin=5.0):
    products=d["products"].copy(); batches=d["batches"].copy(); out=[]
    for r in products.itertuples():
        stock=float(r.stock_qty); daily=float(r.daily_sales_rate); cover=float(r.stock_cover_days)
        safe=max_safe_discount(float(r.unit_price_EGP),float(r.unit_cost_EGP),min_margin)
        exp=batches[batches.product_id==r.product_id]
        expired_qty=float(exp.loc[exp.days_left<0,"batch_qty_units"].sum())
        upcoming=exp[exp.days_left.between(0,45,inclusive="both")]
        exp_qty=float(upcoming["batch_qty_units"].sum())
        min_days=upcoming["days_left"].min() if not upcoming.empty else None
        if expired_qty>0:
            out.append({"priority":0,"type":"دفعات منتهية","product":r.product_name,"product_id":r.product_id,
                        "evidence":f"{expired_qty:,.0f} وحدة منتهية الصلاحية موجودة ضمن الدفعات المسجلة",
                        "action":"اعزل الدفعة فورًا ولا تعرضها للبيع؛ راجع المورد وإجراءات المرتجعات والتالف",
                        "discount":0,"qty":expired_qty,"status":"إجراء فوري — لا تُباع","safe":0})
        if exp_qty>0:
            urgency="حرج" if min_days<=15 else ("مرتفع" if min_days<=30 else "تنبيه")
            target=15 if min_days<=15 else (10 if min_days<=30 else 5)
            discount=min(target,safe)
            action=f"اعرض كمية {exp_qty:,.0f} وحدة بعرض تصريف بخصم يصل إلى {discount:.1f}% قبل انتهاء الصلاحية"
            out.append({"priority":0 if min_days<=15 else 1,"type":"صلاحية","product":r.product_name,"product_id":r.product_id,"evidence":f"{exp_qty:,.0f} وحدة تنتهي خلال 45 يومًا؛ أقرب دفعة بعد {int(min_days)} يومًا","action":action,"discount":discount,"qty":exp_qty,"status":urgency,"safe":safe})
        elif stock>0 and (daily==0 or cover>90):
            q=min(stock,max(1,round(stock*.25)))
            target=8 if daily==0 else 5; discount=min(target,safe)
            reason="لا توجد مبيعات خلال آخر 60 يومًا" if daily==0 else f"تغطية مخزون تقارب {cover:.0f} يومًا"
            out.append({"priority":2,"type":"مخزون راكد","product":r.product_name,"product_id":r.product_id,"evidence":f"{reason}؛ المخزون {stock:,.0f} وحدة","action":f"اختبر عرضًا على {q:,.0f} وحدة بخصم لا يتجاوز {discount:.1f}% لتحريك المخزون","discount":discount,"qty":q,"status":"فرصة تحسين","safe":safe})
        elif bool(r.reorder_flag):
            gap=max(0,float(r.reorder_point_units)-stock)
            out.append({"priority":3,"type":"إعادة طلب","product":r.product_name,"product_id":r.product_id,"evidence":f"المخزون {stock:,.0f} أقل من حد إعادة الطلب {r.reorder_point_units:,.0f}","action":f"راجع طلب شراء بحوالي {max(gap,float(r.reorder_qty_units)*.5):,.0f} وحدة بعد مراجعة الطلبات المفتوحة","discount":0,"qty":gap,"status":"يحتاج متابعة","safe":safe})
        if stock>0 and r.sold_60d>0 and cover>60 and not exp_qty:
            out.append({"priority":4,"type":"زيادة مخزون","product":r.product_name,"product_id":r.product_id,"evidence":f"مخزون {stock:,.0f} وحدة مقابل تغطية {cover:.0f} يومًا","action":"خفّض أو أوقف إعادة الطلب مؤقتًا وراجع توزيع الكميات بين الفروع","discount":0,"qty":stock,"status":"مراجعة","safe":safe})
    return sorted(out,key=lambda x:(x["priority"],x["product"]))

def quantity_offer(price,cost,qty,stock,days_left,min_margin,tiers):
    """Suggest a quantity discount constrained by a gross-margin floor."""
    safe=max_safe_discount(price,cost,min_margin)
    # Base stock/expiry pressure, then incremental quantity incentive.
    pressure=0
    if days_left is not None and days_left<=15: pressure=8
    elif days_left is not None and days_left<=30: pressure=5
    elif stock>0 and stock/max(qty,1)>90: pressure=3
    tier_discount=0
    for threshold,disc in tiers:
        if qty>=threshold: tier_discount=disc
    suggested=min(safe,pressure+tier_discount)
    discounted=price*(1-suggested/100)
    gross_profit=(discounted-cost)*qty
    return {"safe_cap":safe,"suggested":suggested,"unit_after":discounted,"total":discounted*qty,"gross_profit":gross_profit,"profit_per_unit":discounted-cost}

# ---------------- Sidebar / data ----------------
st.sidebar.markdown("# 🧠 Smart Merchant")
st.sidebar.caption("نظام دعم القرار الذكي | العسال")
source=None
uploaded=st.sidebar.file_uploader("تحميل ملف بيانات Excel",type=["xlsx"])
if uploaded is not None: source=BytesIO(uploaded.getvalue())
elif DEFAULT_FILE.exists(): source=DEFAULT_FILE
else: st.sidebar.warning("ارفع ملف Excel للبدء")
if source is None:
    st.title("Smart Merchant | العسال")
    st.info("ضع ملف Alassal_Wholesale_Synthetic_Data.xlsx بجوار app.py أو ارفعه من القائمة الجانبية.")
    st.stop()
try: tables=read_workbook(source)
except Exception as e:
    st.error(f"تعذر قراءة ملف البيانات: {e}"); st.stop()

with st.sidebar:
    st.markdown("---")
    as_of=st.date_input("تاريخ تقييم المخزون والتنبيهات",value=date.today())
    min_margin=st.slider("أقل هامش ربح مسموح بعد الخصم",min_value=0,max_value=30,value=5,step=1,help="يستخدمه النظام كحاجز أمان عند اقتراح الخصم.")
    page=st.radio("التنقل داخل النظام",["الرئيسية","ذكاء المخزون والصلاحية","التوصيات الذكية","العروض وخصومات الكمية","المبيعات والربحية","العملاء والتحصيل","الموردون والمرتجعات"],index=0)
    st.markdown("---")
    st.caption("نسخة عرض تجريبية — البيانات الحالية اصطناعية وليست سجلات فعلية للشركة.")

d=derive(tables,as_of)
products=d["products"]; batches=d["batches"]; ar=d["ar"]
recs=recommendation_engine(d,tables,min_margin)

st.markdown(f"""<div class="hero"><h1>العسال | Smart Merchant</h1><p>من متابعة المخزون إلى قرارات بيع وشراء وتسعير أكثر ذكاءً — كل توصية مرتبطة بدليل رقمي وإجراء واضح.</p></div>""",unsafe_allow_html=True)

# ---------------- Home ----------------
if page=="الرئيسية":
    st.markdown('<div class="section-title">ملخص تنفيذي لصاحب المحل</div>',unsafe_allow_html=True)
    overdue=float(ar.loc[ar.is_overdue,"balance_EGP"].sum())
    exp_qty=float(batches.loc[batches.days_left.between(0,45,inclusive="both"),"batch_qty_units"].sum())
    dead=int(((products.stock_qty>0)&((products.sold_60d==0)|(products.stock_cover_days>90))).sum())
    low=int(products.reorder_flag.sum())
    c=st.columns(5)
    c[0].metric("صافي المبيعات المسجلة",money(d["net_sales"]))
    c[1].metric("قيمة الخصومات",money(d["discount_value"]))
    c[2].metric("مخزون مهدد خلال 45 يوم",f"{exp_qty:,.0f} وحدة")
    c[3].metric("منتجات راكدة/بطيئة",f"{dead} منتج")
    c[4].metric("مستحقات متأخرة",money(overdue))
    st.caption(f"نافذة تحليل حركة البيع: آخر 60 يومًا من تاريخ آخر فاتورة في الملف ({pd.Timestamp(d['anchor_date']).date()}). تاريخ تقييم الصلاحية: {as_of}.")
    l,r=st.columns([1.15,1])
    with l:
        st.markdown("### أهم ما يحتاج انتباهك")
        for item in recs[:5]:
            cls="danger" if item["priority"]==0 else ("warn" if item["priority"]<=2 else "")
            st.markdown(f"<div class='recommend {cls}'><b>{item['product']}</b> — {item['type']}<br><span>{item['evidence']}</span><br><b>الإجراء:</b> {item['action']}</div>",unsafe_allow_html=True)
        if not recs: st.success("لا توجد توصيات عاجلة وفق حدود التحليل الحالية.")
    with r:
        st.markdown("### توزيع قيمة المخزون بالتكلفة")
        top=products.nlargest(8,"stock_value_cost")
        fig=px.bar(top,x="stock_value_cost",y="product_name",orientation="h",labels={"stock_value_cost":"قيمة المخزون (ج.م)","product_name":"المنتج"})
        fig.update_layout(height=340, yaxis={"categoryorder":"total ascending"},margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig,use_container_width=True)
        st.markdown("### حالات المتابعة")
        status=pd.DataFrame({"الحالة":["دفعات تنتهي خلال 45 يومًا","منتجات تحت حد إعادة الطلب","فواتير عليها رصيد متأخر"],"العدد/الكمية":[f"{exp_qty:,.0f} وحدة",str(low),str(int((ar.balance_EGP>0).sum()))]})
        st.dataframe(status,use_container_width=True,hide_index=True)

# ---------------- Inventory ----------------
elif page=="ذكاء المخزون والصلاحية":
    st.markdown('<div class="section-title">ذكاء المخزون والتنبيهات الذكية</div>',unsafe_allow_html=True)
    exp=batches[batches.days_left.notna() & (batches.days_left<=45)].copy()
    exp=exp[exp.days_left>=0].sort_values("days_left")
    expired=batches[batches.days_left.notna() & (batches.days_left<0)]
    a,b,c,e=st.columns(4)
    a.metric("دفعات منتهية",f"{expired.batch_qty_units.sum():,.0f} وحدة")
    b.metric("تنتهي خلال 15 يوم",f"{exp.loc[exp.days_left<=15,'batch_qty_units'].sum():,.0f} وحدة")
    c.metric("تنتهي خلال 30 يوم",f"{exp.loc[exp.days_left<=30,'batch_qty_units'].sum():,.0f} وحدة")
    e.metric("تنتهي خلال 45 يوم",f"{exp.batch_qty_units.sum():,.0f} وحدة")
    if not exp.empty:
        exp["مستوى الخطورة"]=exp.days_left.apply(lambda x:"حرج" if x<=15 else ("مرتفع" if x<=30 else "تنبيه"))
        view=exp.rename(columns={"product_name":"اسم المنتج","location_name":"المخزن/الفرع","batch_id":"رقم الدفعة","batch_qty_units":"الكمية","expiry_date":"تاريخ الصلاحية","days_left":"الأيام المتبقية"})
        st.markdown("### المنتجات المهددة — الاسم والموقع والكمية")
        st.dataframe(view[["اسم المنتج","المخزن/الفرع","رقم الدفعة","الكمية","تاريخ الصلاحية","الأيام المتبقية","مستوى الخطورة"]],use_container_width=True,hide_index=True)
    else: st.success("لا توجد دفعات ستنتهي خلال 45 يومًا وفق تاريخ التقييم الحالي.")
    st.markdown("### حالة كل منتج")
    inv_view=products[["product_name","category","stock_qty","reorder_point_units","sold_60d","stock_cover_days","stock_value_cost"]].copy()
    inv_view["الحالة"]=inv_view.apply(lambda x:"إعادة طلب" if x.stock_qty<x.reorder_point_units else ("راكد/بطيء" if x.stock_qty>0 and (x.sold_60d==0 or x.stock_cover_days>90) else "طبيعي"),axis=1)
    inv_view=inv_view.rename(columns={"product_name":"المنتج","category":"الفئة","stock_qty":"المخزون","reorder_point_units":"حد إعادة الطلب","sold_60d":"مبيعات 60 يوم","stock_cover_days":"تغطية المخزون بالأيام","stock_value_cost":"قيمة المخزون بالتكلفة"})
    st.dataframe(inv_view.sort_values("المخزون",ascending=False),use_container_width=True,hide_index=True)

# ---------------- Recommendations ----------------
elif page=="التوصيات الذكية":
    st.markdown('<div class="section-title">مركز التوصيات الذكية</div>',unsafe_allow_html=True)
    st.write("محرك قواعد يربط الصلاحية والمخزون وحركة البيع وحدود الربحية. كل توصية تعرض سببها والبيانات التي بُنيت عليها.")
    f=st.multiselect("تصفية نوع التوصية",options=sorted(set(x["type"] for x in recs)),default=sorted(set(x["type"] for x in recs)))
    selected=[x for x in recs if x["type"] in f]
    for i,x in enumerate(selected,1):
        tone="danger" if x["priority"]==0 else ("warn" if x["priority"]<=2 else "")
        st.markdown(f"<div class='recommend {tone}'><b>{i}. {x['product']}</b> <span style='opacity:.75'>| {x['type']} | {x['status']}</span><br><b>الدليل:</b> {x['evidence']}<br><b>التوصية:</b> {x['action']}<br><span class='small-note'>سقف الخصم الآمن وفق هامش الربح المحدد: {x['safe']:.1f}%</span></div>",unsafe_allow_html=True)
    if not selected: st.info("لا توجد توصيات ضمن عوامل التصفية المختارة.")
    st.markdown("### قواعد اتخاذ القرار")
    st.write("- الصلاحية: تصنيف الدفعات إلى 0–15 يومًا (حرج)، 16–30 (مرتفع)، 31–45 (تنبيه).")
    st.write("- الركود: عدم وجود مبيعات في آخر 60 يومًا أو تغطية مخزون تتجاوز 90 يومًا.")
    st.write("- إعادة الطلب: مقارنة رصيد المخزون بحد إعادة الطلب المسجل لكل منتج.")
    st.write("- الخصومات: لا تتجاوز الحد الذي يحافظ على هامش الربح الأدنى الذي يحدده المستخدم.")

# ---------------- Offers ----------------
elif page=="العروض وخصومات الكمية":
    st.markdown('<div class="section-title">منشئ العروض وخصومات الكمية</div>',unsafe_allow_html=True)
    st.write("احسب خصمًا مقترحًا على حسب كمية الشراء، مع حماية هامش الربح. النسب التالية قابلة للتعديل وتظل خاضعة لسقف الخصم الآمن.")
    names=products.sort_values("product_name").product_name.tolist()
    chosen=st.selectbox("اختر المنتج",names)
    row=products[products.product_name==chosen].iloc[0]
    batch_for=batches[batches.product_id==row.product_id]
    future_batches=batch_for[batch_for.days_left.notna() & (batch_for.days_left>=0)] if not batch_for.empty else batch_for
    exp_days=future_batches["days_left"].min() if not future_batches.empty else None
    stock=float(row.stock_qty)
    x1,x2,x3=st.columns(3)
    x1.metric("سعر البيع",money(row.unit_price_EGP)); x2.metric("تكلفة الوحدة",money(row.unit_cost_EGP)); x3.metric("المخزون المتاح",f"{stock:,.0f} {row.unit}")
    st.markdown("#### شرائح خصم الكمية")
    t1,t2,t3=st.columns(3)
    q1=t1.number_input("من كمية",min_value=1,value=5,step=1,key="q1"); d1=t1.number_input("خصم %",min_value=0.0,max_value=50.0,value=2.0,step=.5,key="d1")
    q2=t2.number_input("من كمية",min_value=1,value=10,step=1,key="q2"); d2=t2.number_input("خصم %",min_value=0.0,max_value=50.0,value=4.0,step=.5,key="d2")
    q3=t3.number_input("من كمية",min_value=1,value=20,step=1,key="q3"); d3=t3.number_input("خصم %",min_value=0.0,max_value=50.0,value=6.0,step=.5,key="d3")
    qty=st.number_input("كمية الشراء المراد اختبارها",min_value=1,value=10,step=1)
    tiers=sorted([(int(q1),float(d1)),(int(q2),float(d2)),(int(q3),float(d3))],key=lambda z:z[0])
    result=quantity_offer(float(row.unit_price_EGP),float(row.unit_cost_EGP),int(qty),stock,exp_days,min_margin,tiers)
    st.markdown("#### نتيجة العرض")
    m=st.columns(4)
    m[0].metric("الخصم المقترح",pct(result["suggested"]))
    m[1].metric("سعر الوحدة بعد الخصم",money(result["unit_after"]))
    m[2].metric("إجمالي الفاتورة",money(result["total"]))
    m[3].metric("ربح إجمالي تقديري",money(result["gross_profit"]))
    st.caption(f"سقف الخصم الآمن لهذا المنتج: {result['safe_cap']:.1f}% — الحد الأدنى للهامش المطلوب: {min_margin:.0f}%.")
    if result["suggested"]+0.001 < max([disc for threshold,disc in tiers if qty>=threshold] or [0]):
        st.warning("تم تخفيض الخصم عن الشريحة المطلوبة حتى لا ينخفض هامش الربح عن الحد الأدنى المحدد.")
    if qty>stock: st.warning("الكمية المطلوبة أكبر من المخزون المسجل. راجع التوفر قبل اعتماد العرض.")
    st.markdown("#### جدول مقترح لشرائح الكمية")
    tier_rows=[]
    for threshold,disc in tiers:
        rr=quantity_offer(float(row.unit_price_EGP),float(row.unit_cost_EGP),threshold,stock,exp_days,min_margin,[(threshold,disc)])
        tier_rows.append({"من كمية":threshold,"الخصم المطلوب %":disc,"الخصم الآمن المقترح %":round(rr['suggested'],2),"سعر الوحدة بعد الخصم":round(rr['unit_after'],2),"إجمالي البيع":round(rr['total'],2),"ربح إجمالي تقديري":round(rr['gross_profit'],2)})
    st.dataframe(pd.DataFrame(tier_rows),use_container_width=True,hide_index=True)
    st.info("النظام يقترح ولا يطبق الخصم تلقائيًا. اعتماد النسب النهائية يكون من صاحب المحل بعد مراجعة سياسة التسعير وتكلفة الشراء الفعلية.")

# ---------------- Sales ----------------
elif page=="المبيعات والربحية":
    st.markdown('<div class="section-title">ذكاء المبيعات والربحية</div>',unsafe_allow_html=True)
    si=tables["Sales_Invoice_Items"].copy(); si=si.merge(tables["Products"][["product_id","product_name","category","unit_cost_EGP"]],on="product_id",how="left")
    si["estimated_cost"]=si.quantity*si.unit_cost_EGP
    si["estimated_profit"]=si.line_net_EGP-si.estimated_cost
    si["margin_pct"]=si.apply(lambda x:x.estimated_profit/x.line_net_EGP*100 if x.line_net_EGP else 0,axis=1)
    a,b,c,e=st.columns(4)
    a.metric("إجمالي قبل الخصم",money(d["gross_sales"]))
    b.metric("صافي بعد الخصم",money(d["net_sales"]))
    c.metric("إجمالي الخصومات",money(d["discount_value"]))
    e.metric("متوسط نسبة الخصم",pct((d["discount_value"]/d["gross_sales"]*100) if d["gross_sales"] else 0))
    by=si.groupby("product_name",as_index=False).agg(net_sales=("line_net_EGP","sum"),estimated_profit=("estimated_profit","sum"),qty=("quantity","sum"),discount=("discount_pct","mean")).sort_values("net_sales",ascending=False).head(12)
    l,r=st.columns(2)
    with l:
        st.markdown("### أعلى المنتجات مبيعًا")
        fig=px.bar(by.sort_values("net_sales"),x="net_sales",y="product_name",orientation="h",labels={"net_sales":"صافي المبيعات (ج.م)","product_name":"المنتج"})
        fig.update_layout(height=430,margin=dict(l=10,r=10,t=10,b=10)); st.plotly_chart(fig,use_container_width=True)
    with r:
        st.markdown("### الربح الإجمالي التقديري حسب المنتج")
        pr=by.sort_values("estimated_profit",ascending=False).head(12)
        fig=px.bar(pr.sort_values("estimated_profit"),x="estimated_profit",y="product_name",orientation="h",labels={"estimated_profit":"ربح تقديري (ج.م)","product_name":"المنتج"})
        fig.update_layout(height=430,margin=dict(l=10,r=10,t=10,b=10)); st.plotly_chart(fig,use_container_width=True)
    st.markdown("### تفاصيل المنتجات")
    st.dataframe(by.rename(columns={"product_name":"المنتج","net_sales":"صافي المبيعات","estimated_profit":"الربح التقديري","qty":"الكمية المباعة","discount":"متوسط الخصم %"}),use_container_width=True,hide_index=True)
    st.caption("الربح تقديري باستخدام تكلفة الوحدة الحالية في ورقة Products، وليس تكلفة تاريخية موزونة. يجب اعتماد طريقة COGS مع الشركة قبل استخدامه محاسبيًا.")

# ---------------- Customers ----------------
elif page=="العملاء والتحصيل":
    st.markdown('<div class="section-title">ذكاء العملاء والتحصيل</div>',unsafe_allow_html=True)
    overdue=ar[ar.is_overdue & (ar.balance_EGP>0)].copy().sort_values("balance_EGP",ascending=False)
    a,b,c=st.columns(3)
    a.metric("إجمالي الأرصدة المستحقة",money(ar.balance_EGP.sum()))
    b.metric("أرصدة متأخرة",money(overdue.balance_EGP.sum()))
    c.metric("عدد الفواتير المتأخرة",f"{len(overdue):,}")
    if not overdue.empty:
        overdue["days_overdue"]=overdue.days_overdue.fillna(0).clip(lower=0)
        st.markdown("### العملاء والفواتير المطلوب متابعتها")
        st.dataframe(overdue[["customer_name","invoice_id","order_date","due_date","total_amount_EGP","paid_EGP","balance_EGP","days_overdue"]].rename(columns={"customer_name":"العميل","invoice_id":"رقم الفاتورة","order_date":"تاريخ الفاتورة","due_date":"تاريخ الاستحقاق","total_amount_EGP":"قيمة الفاتورة","paid_EGP":"المدفوع","balance_EGP":"الرصيد المتبقي","days_overdue":"أيام التأخير"}),use_container_width=True,hide_index=True)
    else: st.success("لا توجد أرصدة متأخرة وفق تاريخ التقييم.")
    customer_ar=ar.groupby("customer_name",as_index=False).agg(invoices=("invoice_id","count"),balance=("balance_EGP","sum"),sales=("total_amount_EGP","sum")).sort_values("balance",ascending=False).head(15)
    st.markdown("### أعلى العملاء حسب الرصيد المتبقي")
    st.dataframe(customer_ar.rename(columns={"customer_name":"العميل","invoices":"عدد الفواتير","balance":"الرصيد المتبقي","sales":"إجمالي الفواتير"}),use_container_width=True,hide_index=True)

# ---------------- Suppliers and returns ----------------
elif page=="الموردون والمرتجعات":
    st.markdown('<div class="section-title">أداء الموردين والمرتجعات</div>',unsafe_allow_html=True)
    po=tables["Purchase_Orders"].merge(tables["Suppliers"][["supplier_id","supplier_name","avg_lead_time_days","quality_rating_1to5"]],on="supplier_id",how="left")
    po["actual_delivery_date"]=pd.to_datetime(po.actual_delivery_date,errors="coerce"); po["expected_delivery_date"]=pd.to_datetime(po.expected_delivery_date,errors="coerce")
    po["delay_days"]=(po.actual_delivery_date-po.expected_delivery_date).dt.days
    delivered=po[po.actual_delivery_date.notna()]
    supplier=delivered.groupby("supplier_name",as_index=False).agg(orders=("po_id","count"),avg_delay=("delay_days","mean"),quality=("quality_rating_1to5","mean")).sort_values("avg_delay",ascending=False)
    a,b,c=st.columns(3)
    a.metric("عدد أوامر الشراء",f"{len(po):,}")
    b.metric("متوسط تأخير التسليم",f"{delivered.delay_days.mean():.1f} يوم" if not delivered.empty else "—")
    c.metric("عدد المرتجعات",f"{len(tables['Returns']):,}")
    st.markdown("### مؤشرات الموردين")
    st.dataframe(supplier.rename(columns={"supplier_name":"المورد","orders":"عدد الطلبات","avg_delay":"متوسط الفرق عن موعد التسليم (يوم)","quality":"تقييم الجودة"}),use_container_width=True,hide_index=True)
    ret=tables["Returns"].merge(tables["Products"][["product_id","product_name"]],on="product_id",how="left")
    ret_summary=ret.groupby(["product_name","return_reason"],as_index=False).agg(return_qty=("quantity_returned","sum"),events=("return_id","count")).sort_values("return_qty",ascending=False)
    st.markdown("### المنتجات وأسباب المرتجعات")
    st.dataframe(ret_summary.rename(columns={"product_name":"المنتج","return_reason":"سبب الإرجاع","return_qty":"كمية مرتجعة","events":"عدد الحالات"}),use_container_width=True,hide_index=True)

st.markdown("---")
st.caption("Smart Merchant Demo | جميع المؤشرات والتوصيات تعتمد على ملف البيانات المرفق. البيانات الحالية اصطناعية؛ يلزم اختبارها على عينة فعلية ومراجعة قواعد التشغيل مع صاحب المحل قبل اعتمادها.")
