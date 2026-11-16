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
    paid=pay.groupby("invoice_id",as_index=False)["amount_paid_EGP"].sum().rename(columns={"amount_paid_EGP":"paid_EGP"})
    ar=inv.merge(paid,on="invoice_id",how="left").merge(cust[["customer_id","customer_name","customer_type"]],on="customer_id",how="left")
    ar["paid_EGP"]=ar["paid_EGP"].fillna(0)
    ar["balance_EGP"]=(ar["total_amount_EGP"]-ar["paid_EGP"]).clip(lower=0)
    ar["due_date"]=pd.to_datetime(ar["due_date"],errors="coerce")
    ar["days_overdue"]=(pd.Timestamp(as_of)-ar["due_date"]).dt.days
    ar["is_overdue"]=(ar.balance_EGP>0)&(ar.due_date.dt.date<as_of)
    gross=si["line_gross_EGP"].sum(); net=si["line_net_EGP"].sum()
    return {"products":prod,"batches":b,"sales":sales,"recent_sales":recent,"ar":ar,"gross_sales":gross,"net_sales":net,
            "discount_value":gross-net,"anchor_date":anchor,"window_start":window_start}


def filter_tables_by_location(tables, selected_location):
    """Return branch-scoped tables; stock follows location_id and sales follow branch_id."""
    scoped = {name: frame.copy() for name, frame in tables.items()}
    if selected_location == "ALL":
        return scoped

    invoices = scoped["Sales_Invoices"]
    invoices = invoices[invoices["branch_id"].astype(str) == str(selected_location)].copy()
    invoice_ids = set(invoices["invoice_id"].astype(str))
    scoped["Sales_Invoices"] = invoices
    scoped["Sales_Invoice_Items"] = scoped["Sales_Invoice_Items"][
        scoped["Sales_Invoice_Items"]["invoice_id"].astype(str).isin(invoice_ids)
    ].copy()
    scoped["Payments_Collections"] = scoped["Payments_Collections"][
        scoped["Payments_Collections"]["invoice_id"].astype(str).isin(invoice_ids)
    ].copy()
    scoped["Returns"] = scoped["Returns"][
        scoped["Returns"]["invoice_id"].astype(str).isin(invoice_ids)
    ].copy()
    scoped["Inventory_Batches"] = scoped["Inventory_Batches"][
        scoped["Inventory_Batches"]["location_id"].astype(str) == str(selected_location)
    ].copy()
    scoped["Customers"] = scoped["Customers"][
        scoped["Customers"]["branch_id"].astype(str) == str(selected_location)
    ].copy()
    return scoped


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


def customer_analytics(tables, d):
    """Explainable customer metrics; small demo data is not enough for supervised churn training."""
    invoices = tables["Sales_Invoices"].copy()
    customers = tables["Customers"].copy()
    invoices["order_date"] = pd.to_datetime(invoices["order_date"], errors="coerce")
    anchor = pd.to_datetime(d["anchor_date"]).normalize()
    base = invoices.merge(customers[["customer_id", "customer_name", "customer_type", "region"]], on="customer_id", how="right")
    rows = []
    for customer_id, g in base.groupby("customer_id", dropna=False):
        g = g.dropna(subset=["order_date"]).sort_values("order_date")
        count = int(g["invoice_id"].nunique())
        total = float(g["total_amount_EGP"].sum()) if not g.empty else 0.0
        first = g["order_date"].min() if not g.empty else pd.NaT
        last = g["order_date"].max() if not g.empty else pd.NaT
        recency = int((anchor - last).days) if pd.notna(last) else 999
        gaps = g["order_date"].diff().dt.days.dropna()
        avg_gap = float(gaps.mean()) if not gaps.empty else 45.0
        active_months = max(1.0, ((anchor - first).days / 30.4375 + 1) if pd.notna(first) else 1.0)
        monthly_sales = total / active_months
        expected_gap = max(30.0, avg_gap)
        churn_score = min(100.0, max(0.0, recency / (expected_gap * 1.5) * 100))
        if count == 0:
            segment = "لا يوجد سجل شراء"
        elif churn_score >= 75:
            segment = "معرض للخمول"
        elif count == 1 and recency <= 45:
            segment = "عميل جديد"
        elif monthly_sales >= 1 and count >= 3 and monthly_sales >= float(base.groupby("customer_id")["total_amount_EGP"].sum().median() if not base.empty else 0):
            segment = "عميل ذو قيمة عالية"
        elif count >= 3 and churn_score < 45:
            segment = "عميل منتظم"
        elif churn_score >= 45:
            segment = "يحتاج متابعة"
        else:
            segment = "عميل نشط"
        expected_months = max(1.0, 12.0 * (1.0 - churn_score / 100.0))
        projected_clv = monthly_sales * expected_months
        customer_name = customers.loc[customers.customer_id == customer_id, "customer_name"].iloc[0]
        customer_type = customers.loc[customers.customer_id == customer_id, "customer_type"].iloc[0]
        region = customers.loc[customers.customer_id == customer_id, "region"].iloc[0]
        rows.append({
            "customer_id": customer_id, "customer_name": customer_name, "customer_type": customer_type,
            "region": region, "invoice_count": count, "historical_sales": total,
            "monthly_sales_avg": monthly_sales, "last_purchase": last,
            "days_since_purchase": recency, "avg_purchase_gap_days": avg_gap,
            "churn_risk_index": churn_score, "segment": segment,
            "estimated_active_months": expected_months, "estimated_clv": projected_clv,
        })
    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame(columns=["customer_id", "customer_name", "customer_type", "region", "invoice_count", "historical_sales", "monthly_sales_avg", "last_purchase", "days_since_purchase", "avg_purchase_gap_days", "churn_risk_index", "segment", "estimated_active_months", "estimated_clv"])
    return result.sort_values(["churn_risk_index", "historical_sales"], ascending=[False, False])


def inventory_purchase_plan(tables, d, as_of, coverage_months, safety_pct=10):
    """Calculate purchasing needs for the whole shop, using all customers' observed product demand."""
    prod = tables["Products"].copy()
    sales = d["sales"].copy()
    batches = d["batches"].copy()
    anchor = pd.to_datetime(d["anchor_date"]).normalize()
    window_start = anchor - pd.Timedelta(days=59)
    recent = sales[sales["order_date"] >= window_start]
    demand = recent.groupby("product_id", as_index=False).agg(sold_60d=("quantity", "sum"))
    demand["avg_monthly_demand"] = demand["sold_60d"] / 2.0

    target_date = pd.Timestamp(as_of) + pd.DateOffset(months=int(coverage_months))
    batches["expiry_date"] = pd.to_datetime(batches["expiry_date"], errors="coerce")
    valid = batches[(batches["expiry_date"].isna()) | (batches["expiry_date"] >= target_date)]
    usable_stock = valid.groupby("product_id")["batch_qty_units"].sum()
    total_stock = batches.groupby("product_id")["batch_qty_units"].sum()
    expiring_before_target = batches[
        batches["expiry_date"].notna() & (batches["expiry_date"] < target_date) & (batches["expiry_date"] >= pd.Timestamp(as_of))
    ].groupby("product_id")["batch_qty_units"].sum()

    plan = prod.merge(demand[["product_id", "sold_60d", "avg_monthly_demand"]], on="product_id", how="left")
    plan["sold_60d"] = plan["sold_60d"].fillna(0)
    plan["avg_monthly_demand"] = plan["avg_monthly_demand"].fillna(0)
    plan["valid_stock"] = plan["product_id"].map(usable_stock).fillna(0)
    plan["current_stock"] = plan["product_id"].map(total_stock).fillna(0)
    plan["expiring_before_horizon"] = plan["product_id"].map(expiring_before_target).fillna(0)
    plan["forecast_need"] = plan["avg_monthly_demand"] * int(coverage_months)
    plan["safety_stock"] = plan["forecast_need"] * (float(safety_pct) / 100.0)
    plan["target_stock"] = (plan["forecast_need"] + plan["safety_stock"]).apply(math.ceil)
    plan["suggested_purchase"] = (plan["target_stock"] - plan["valid_stock"]).clip(lower=0).apply(math.ceil)
    plan["estimated_purchase_cost"] = plan["suggested_purchase"] * plan["unit_cost_EGP"]
    supplier_names = tables["Suppliers"][["supplier_id", "supplier_name"]]
    plan = plan.merge(supplier_names, left_on="primary_supplier_id", right_on="supplier_id", how="left")
    plan["status"] = plan.apply(
        lambda r: "لا توجد مبيعات حديثة" if r.avg_monthly_demand <= 0 else ("يحتاج شراء" if r.suggested_purchase > 0 else "المخزون يغطي الفترة"), axis=1
    )
    return plan.sort_values(["suggested_purchase", "avg_monthly_demand"], ascending=[False, False]), target_date


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
    locations = tables["Branches_Warehouses"][["location_id", "location_name"]].drop_duplicates()
    location_options = {"كل الفروع والمواقع": "ALL"}
    location_options.update({str(r.location_name): str(r.location_id) for r in locations.itertuples()})
    selected_location_name = st.selectbox("اختيار الفرع / الموقع", list(location_options.keys()), index=0)
    selected_location = location_options[selected_location_name]
    st.caption("يشمل الاختيار الفروع والمخزن الرئيسي. تتغير المبيعات حسب الفرع، والمخزون حسب موقع التخزين.")
    as_of=st.date_input("تاريخ تقييم المخزون والتنبيهات",value=date.today())
    min_margin=st.slider("أقل هامش ربح مسموح بعد الخصم",min_value=0,max_value=30,value=5,step=1,help="يستخدمه النظام كحاجز أمان عند اقتراح الخصم.")
    page=st.radio("التنقل داخل النظام",["الرئيسية","ذكاء المخزون والصلاحية","التخطيط الذكي للمخزون","التوصيات الذكية","العروض وخصومات الكمية","المبيعات والربحية","العملاء والتحصيل","الموردون والمرتجعات"],index=0)
    st.markdown("---")
    st.caption("نسخة عرض تجريبية — البيانات الحالية اصطناعية وليست سجلات فعلية للشركة.")

all_tables = tables
tables = filter_tables_by_location(all_tables, selected_location)
d=derive(tables,as_of)
products=d["products"]; batches=d["batches"]; ar=d["ar"]
recs=recommendation_engine(d,tables,min_margin)

st.markdown(f"""<div class="hero"><h1>العسال | Smart Merchant</h1><p>من متابعة المخزون إلى قرارات بيع وشراء وتسعير أكثر ذكاءً — كل توصية مرتبطة بدليل رقمي وإجراء واضح.</p></div>""",unsafe_allow_html=True)
st.info(f"الموقع المحدد حاليًا: **{selected_location_name}** — جميع المؤشرات المعروضة مفلترة وفق هذا الاختيار.")

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


# ---------------- Smart inventory purchasing plan ----------------
elif page=="التخطيط الذكي للمخزون":
    st.markdown('<div class="section-title">التخطيط الذكي للمخزون والمشتريات</div>',unsafe_allow_html=True)
    st.write("حدد مدة التغطية المطلوبة، وسيحسب النظام احتياجات المحل من جميع العملاء اعتمادًا على حركة بيع المنتجات، ثم يخصم المخزون الصالح المتاح خلال مدة التغطية.")
    x1, x2 = st.columns(2)
    with x1:
        coverage = st.slider("تغطية المخزون المطلوبة (بالشهور)", min_value=1, max_value=12, value=3, step=1)
    with x2:
        safety = st.slider("نسبة مخزون الأمان الإضافية", min_value=0, max_value=30, value=10, step=5, help="هامش إضافي اختياري لمواجهة زيادة الطلب.")
    plan, target_date = inventory_purchase_plan(tables, d, as_of, coverage, safety)
    buy = plan[plan.suggested_purchase > 0].copy()
    total_cost = float(buy.estimated_purchase_cost.sum())
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("مدة التغطية", f"{coverage} شهر")
    k2.metric("عدد المنتجات المطلوب شراؤها", f"{len(buy):,} منتج")
    k3.metric("إجمالي الوحدات المقترحة", f"{buy.suggested_purchase.sum():,.0f}")
    k4.metric("تكلفة الشراء التقديرية", money(total_cost))
    st.caption(f"تاريخ نهاية التغطية: {target_date.date()} | الطلب محسوب من آخر 60 يومًا من حركة الفواتير المسجلة حتى {pd.Timestamp(d['anchor_date']).date()}.")
    st.markdown("### قائمة المنتجات والكميات المقترحة للشراء")
    if not buy.empty:
        view = buy[["product_name", "category", "unit", "avg_monthly_demand", "forecast_need", "valid_stock", "expiring_before_horizon", "suggested_purchase", "unit_cost_EGP", "estimated_purchase_cost", "supplier_name", "status"]].copy()
        view = view.rename(columns={"product_name":"اسم المنتج", "category":"الفئة", "unit":"الوحدة", "avg_monthly_demand":"متوسط الطلب الشهري", "forecast_need":"احتياج فترة التغطية", "valid_stock":"مخزون صالح حتى نهاية الفترة", "expiring_before_horizon":"مخزون تنتهي صلاحيته قبل نهاية الفترة", "suggested_purchase":"الكمية المقترح شراؤها", "unit_cost_EGP":"تكلفة الوحدة", "estimated_purchase_cost":"تكلفة الشراء التقديرية", "supplier_name":"المورد الأساسي", "status":"الحالة"})
        st.dataframe(view, use_container_width=True, hide_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            view.to_excel(writer, index=False, sheet_name="خطة المشتريات")
        st.download_button("تحميل خطة المشتريات Excel", data=output.getvalue(), file_name="Alassal_Inventory_Purchase_Plan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.success("المخزون الصالح الحالي يغطي الطلب المتوقع للفترة المختارة وفق البيانات المتاحة.")
    with st.expander("عرض تحليل جميع المنتجات"):
        all_view = plan[["product_name", "avg_monthly_demand", "forecast_need", "current_stock", "valid_stock", "expiring_before_horizon", "target_stock", "suggested_purchase", "status"]].rename(columns={"product_name":"المنتج", "avg_monthly_demand":"متوسط الطلب الشهري", "forecast_need":"الطلب المتوقع", "current_stock":"إجمالي المخزون الحالي", "valid_stock":"المخزون الصالح للفترة", "expiring_before_horizon":"كمية تنتهي قبل نهاية الفترة", "target_stock":"المخزون المستهدف", "suggested_purchase":"المطلوب شراؤه", "status":"الحالة"})
        st.dataframe(all_view, use_container_width=True, hide_index=True)
    st.info("ملاحظة: الحساب يعتمد على متوسط مبيعات آخر 60 يومًا. الكميات التي تنتهي صلاحيتها قبل نهاية فترة التغطية لا تُحسب ضمن المخزون القادر على تغطية الفترة. هذه خطة اقتراحية وليست أمر شراء تلقائيًا.")


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
    tab_ar, tab_segments, tab_value = st.tabs(["التحصيل والأرصدة", "تقسيم العملاء", "التخلي عن الشراء والقيمة الدائمة"])
    with tab_ar:
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

    customer_kpis = customer_analytics(tables, d)
    with tab_segments:
        st.markdown("### تقسيم العملاء حسب سلوك الشراء")
        if customer_kpis.empty:
            st.info("لا توجد بيانات عملاء مرتبطة بهذا الموقع في ملف البيانات الحالي.")
        seg_counts = customer_kpis.groupby("segment", as_index=False).agg(customers=("customer_id", "count"), sales=("historical_sales", "sum"))
        c1, c2, c3 = st.columns(3)
        c1.metric("عدد العملاء", f"{len(customer_kpis):,}")
        c2.metric("عملاء منتظمون", f"{int(customer_kpis.segment.eq('عميل منتظم').sum()):,}")
        c3.metric("عملاء معرضون للخمول", f"{int(customer_kpis.segment.eq('معرض للخمول').sum()):,}")
        st.dataframe(seg_counts.rename(columns={"segment":"الشريحة", "customers":"عدد العملاء", "sales":"إجمالي المبيعات"}), use_container_width=True, hide_index=True)
        st.markdown("### تفاصيل كل عميل")
        segment_view = customer_kpis[["customer_name", "customer_type", "region", "invoice_count", "historical_sales", "monthly_sales_avg", "days_since_purchase", "segment"]].rename(columns={"customer_name":"العميل", "customer_type":"نوع العميل", "region":"المنطقة", "invoice_count":"عدد الفواتير", "historical_sales":"إجمالي المبيعات التاريخية", "monthly_sales_avg":"متوسط المبيعات الشهرية", "days_since_purchase":"أيام منذ آخر شراء", "segment":"تصنيف العميل"})
        st.dataframe(segment_view, use_container_width=True, hide_index=True)
        st.caption("التقسيم الحالي قائم على حداثة الشراء وتكراره وقيمة مشترياته، وهو تقسيم تفسيري قابل للمراجعة وليس تصنيفًا مدربًا على بيانات موسومة.")

    with tab_value:
        st.markdown("### مؤشر احتمالية التخلي عن الشراء")
        risk = customer_kpis.copy()
        risk["مستوى الخطورة"] = risk.churn_risk_index.apply(lambda x: "مرتفع" if x >= 75 else ("متوسط" if x >= 45 else "منخفض"))
        high_risk = risk[risk.churn_risk_index >= 75]
        a, b, c, e = st.columns(4)
        a.metric("عملاء خطورة مرتفعة", f"{len(high_risk):,}")
        b.metric("نسبة العملاء عاليي الخطورة", pct(len(high_risk) / len(risk) * 100 if len(risk) else 0))
        c.metric("عملاء خطورة متوسطة", f"{int(risk.churn_risk_index.between(45, 74.999).sum()):,}")
        e.metric("متوسط مؤشر الخطورة", pct(risk.churn_risk_index.mean() if len(risk) else 0))
        risk_view = risk[["customer_name", "days_since_purchase", "avg_purchase_gap_days", "invoice_count", "churn_risk_index", "مستوى الخطورة"]].rename(columns={"customer_name":"العميل", "days_since_purchase":"أيام منذ آخر شراء", "avg_purchase_gap_days":"متوسط الفاصل بين المشتريات (يوم)", "invoice_count":"عدد الفواتير", "churn_risk_index":"مؤشر الخطورة %"})
        st.dataframe(risk_view, use_container_width=True, hide_index=True)
        st.markdown("### القيمة الدائمة التقديرية للعملاء")
        clv_total = float(customer_kpis.estimated_clv.sum())
        st.metric("إجمالي القيمة المستقبلية التقديرية للعملاء", money(clv_total))
        clv_view = customer_kpis.sort_values("estimated_clv", ascending=False)[["customer_name", "historical_sales", "monthly_sales_avg", "estimated_active_months", "estimated_clv", "churn_risk_index"]].rename(columns={"customer_name":"العميل", "historical_sales":"المبيعات التاريخية", "monthly_sales_avg":"متوسط المبيعات الشهرية", "estimated_active_months":"مدة استمرار متوقعة (شهر)", "estimated_clv":"القيمة الدائمة التقديرية", "churn_risk_index":"مؤشر التخلي %"})
        st.dataframe(clv_view, use_container_width=True, hide_index=True)
        st.caption("مهم: هذا إصدار تقديري قائم على قواعد سلوكية بسبب محدودية عدد العملاء وعدم وجود سجل تاريخي طويل وموسوم بحالات ترك الخدمة. مؤشر الخطورة ليس احتمالًا إحصائيًا معايرًا، والقيمة الدائمة سيناريو تقديري مبني على متوسط الإنفاق الشهري ومدة استمرار افتراضية تتأثر بمؤشر الخطورة؛ لا يُعرض كناتج نموذج ML موثوق قبل توافر بيانات فعلية أكثر.")

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
    st.caption("مؤشرات أوامر الشراء معروضة على مستوى الشركة؛ ملف البيانات الحالي لا يربط أوامر الشراء بموقع استلام محدد.")
    st.dataframe(supplier.rename(columns={"supplier_name":"المورد","orders":"عدد الطلبات","avg_delay":"متوسط الفرق عن موعد التسليم (يوم)","quality":"تقييم الجودة"}),use_container_width=True,hide_index=True)
    ret=tables["Returns"].merge(tables["Products"][["product_id","product_name"]],on="product_id",how="left")
    ret_summary=ret.groupby(["product_name","return_reason"],as_index=False).agg(return_qty=("quantity_returned","sum"),events=("return_id","count")).sort_values("return_qty",ascending=False)
    st.markdown("### المنتجات وأسباب المرتجعات")
    st.dataframe(ret_summary.rename(columns={"product_name":"المنتج","return_reason":"سبب الإرجاع","return_qty":"كمية مرتجعة","events":"عدد الحالات"}),use_container_width=True,hide_index=True)

st.markdown("---")
st.caption("Smart Merchant Demo | جميع المؤشرات والتوصيات تعتمد على ملف البيانات المرفق. البيانات الحالية اصطناعية؛ يلزم اختبارها على عينة فعلية ومراجعة قواعد التشغيل مع صاحب المحل قبل اعتمادها.")
