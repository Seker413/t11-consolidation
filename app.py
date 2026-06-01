"""
基于AI规则引擎的内部存货交易未实现利润自动抵消系统
====================================================
V2.0 | 模块化架构 | 2025级会计专硕课程作业
"""
import streamlit as st
import pandas as pd
import json
import sqlite3
import os
import io
from datetime import datetime
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.font_manager as fm

# 注册本地中文字体（兼容Windows/Linux/Streamlit Cloud）
_font_path = Path(__file__).parent / "fonts" / "simhei.ttf"
if _font_path.exists():
    fm.fontManager.addfont(str(_font_path))
    _font_name = fm.FontProperties(fname=str(_font_path)).get_name()
    matplotlib.rcParams['font.sans-serif'] = [_font_name, 'DejaVu Sans']
else:
    matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

# ---------- 项目内部模块 ----------
from engine.trade_classifier import (
    build_equity_graph, classify_trade, get_minority_ratio, detect_anomalies
)
from engine.calc_engine import calculate_unrealized_profit, compute_summary
from engine.entry_generator import generate_entries, get_trade_label
from utils.hash_utils import compute_hash
from utils.pdf_export import export_pdf

# ---------- 配置 ----------
st.set_page_config(page_title="内部存货交易未实现利润自动抵消", page_icon="📊", layout="wide")
DB_PATH = Path(__file__).parent / "data" / "calc_log.db"


# ========== 数据库 ==========
def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("""CREATE TABLE IF NOT EXISTS equity_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT, parent_code TEXT NOT NULL,
        child_code TEXT NOT NULL, share_ratio REAL NOT NULL,
        is_consolidated INTEGER DEFAULT 1,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(parent_code, child_code))""")
    conn.execute("""CREATE TABLE IF NOT EXISTS calc_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT, elim_no TEXT NOT NULL,
        batch_id TEXT NOT NULL, trade_json TEXT NOT NULL,
        calc_json TEXT NOT NULL, entries_json TEXT NOT NULL,
        log_hash TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
    conn.commit()
    return conn


def save_equity_to_db(conn, df):
    conn.execute("DELETE FROM equity_config")
    for _, row in df.iterrows():
        conn.execute("INSERT INTO equity_config (parent_code, child_code, share_ratio, is_consolidated) VALUES (?,?,?,?)",
                     (str(row.iloc[0]), str(row.iloc[1]), float(row.iloc[2]),
                      1 if str(row.iloc[3]).strip() in ("是","1","True","true") else 0))
    conn.commit()


def load_equity_from_db(conn):
    return pd.read_sql("SELECT parent_code, child_code, share_ratio, is_consolidated FROM equity_config", conn)


# ========== 校验 ==========
def validate_equity_df(df):
    required = ["母公司编码", "子公司编码", "母公司持股比例", "是否纳入合并"]
    missing = [c for c in required if c not in df.columns]
    if missing: return [f"缺少列：{', '.join(missing)}"]
    if not pd.api.types.is_numeric_dtype(df["母公司持股比例"]): return ["「母公司持股比例」包含非数值"]
    return []

def validate_sales_df(df):
    required = ["销售方编码", "购买方编码", "存货批次号", "存货名称", "销售收入", "销售成本", "销售数量", "会计期间"]
    missing = [c for c in required if c not in df.columns]
    if missing: return [f"缺少列：{', '.join(missing)}"]
    for col in ["销售收入", "销售成本", "销售数量"]:
        if not pd.api.types.is_numeric_dtype(df[col]): return [f"「{col}」包含非数值"]
    return []

def validate_inventory_df(df):
    required = ["主体编码", "存货批次号", "期末结存数量", "内部购入总数量"]
    missing = [c for c in required if c not in df.columns]
    if missing: return [f"缺少列：{', '.join(missing)}"]
    for col in ["期末结存数量", "内部购入总数量"]:
        if not pd.api.types.is_numeric_dtype(df[col]): return [f"「{col}」包含非数值"]
    return []


# ========== 核心引擎 ==========
def run_engine(sales_df, inventory_df, equity_df, tax_rate):
    parent_of, ratio_of, parents, children = build_equity_graph(equity_df)

    # 异常检测
    anomalies, anomaly_count = detect_anomalies(sales_df, inventory_df, equity_df, parent_of)

    merged = sales_df.merge(inventory_df, left_on=["购买方编码","存货批次号"],
                            right_on=["主体编码","存货批次号"], how="left", suffixes=("","_inv"))

    results, logs = [], []
    elim_counter = 0
    period = str(sales_df["会计期间"].iloc[0]) if len(sales_df) > 0 else datetime.now().strftime("%Y%m")
    batch_id = datetime.now().strftime("%Y%m%d%H%M%S")
    prev_hash = "0" * 16
    skipped = 0

    for _, row in merged.iterrows():
        revenue = float(row["销售收入"]); cost = float(row["销售成本"])
        quantity = int(row["销售数量"])
        seller = str(row["销售方编码"]).strip(); buyer = str(row["购买方编码"]).strip()
        batch_no = str(row["存货批次号"]); item_name = str(row["存货名称"])

        ending_qty = int(row["期末结存数量"]) if pd.notna(row.get("期末结存数量")) else 0

        if ending_qty == 0:
            skipped += 1; continue

        trade_type, _ = classify_trade(seller, buyer, parent_of, parents, children)
        minority_ratio = get_minority_ratio(trade_type, seller, buyer, parent_of, ratio_of)

        if trade_type == "DOWN": minority_ratio = None
        elif trade_type in ("UP","FLAT") and minority_ratio is None: minority_ratio = 0.0
        if trade_type == "UNKNOWN": trade_type = "DOWN"

        gmr, unsold, unrealized = calculate_unrealized_profit(revenue, cost, quantity, ending_qty)
        if unrealized <= 0: skipped += 1; continue

        context = {"revenue":revenue,"cost":cost,"unrealized_profit":unrealized,
                   "unsold_ratio":unsold,"tax_rate":tax_rate,"minority_ratio":minority_ratio or 0}
        entries = generate_entries(trade_type, context)

        elim_counter += 1
        elim_no = f"ELIM-{period}-{elim_counter:04d}"

        data_str = json.dumps({"elim_no":elim_no,"seller":seller,"buyer":buyer,
                               "batch_no":batch_no,"trade_type":trade_type,
                               "unrealized_profit":unrealized}, sort_keys=True, ensure_ascii=False)
        log_hash = compute_hash(prev_hash, data_str)
        prev_hash = log_hash

        results.append({"elim_no":elim_no,"trade_type":trade_type,
                        "trade_label":get_trade_label(trade_type),
                        "seller":seller,"buyer":buyer,"batch_no":batch_no,
                        "item_name":item_name,"revenue":revenue,"cost":cost,
                        "quantity":quantity,"ending_qty":ending_qty,
                        "gross_margin_rate":gmr,"unsold_ratio":unsold,
                        "unrealized_profit":unrealized,"minority_ratio":minority_ratio,
                        "entries":entries})

        logs.append({"elim_no":elim_no,"batch_id":batch_id,
                     "trade_json":json.dumps({"seller":seller,"buyer":buyer,"batch_no":batch_no,
                                              "revenue":revenue,"cost":cost,"quantity":quantity,
                                              "ending_qty":ending_qty},ensure_ascii=False),
                     "calc_json":json.dumps({"trade_type":trade_type,"minority_ratio":minority_ratio,
                                             "gross_margin_rate":gmr,"unsold_ratio":unsold,
                                             "unrealized_profit":unrealized,"tax_rate":tax_rate},ensure_ascii=False),
                     "entries_json":json.dumps(entries,ensure_ascii=False),"log_hash":log_hash})

    summary = compute_summary(results, tax_rate, len(sales_df), anomaly_count, skipped)
    return summary, results, logs, batch_id, anomalies


def save_logs_to_db(conn, logs):
    for log in logs:
        conn.execute("INSERT INTO calc_log (elim_no,batch_id,trade_json,calc_json,entries_json,log_hash) VALUES (?,?,?,?,?,?)",
                     (log["elim_no"],log["batch_id"],log["trade_json"],log["calc_json"],log["entries_json"],log["log_hash"]))
    conn.commit()


# ========== Excel导出 ==========
def export_excel(results, summary):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_cn = pd.DataFrame([{
            "内部交易总笔数": summary["total_trades"],
            "需抵消笔数": summary["eliminated"],
            "无需抵消笔数": summary["skipped"],
            "异常笔数": summary.get("anomaly_count", 0),
            "抵消营业收入合计": summary["total_revenue_offset"],
            "抵消营业成本合计": summary["total_cost_offset"],
            "冲减存货合计": summary["total_inventory_offset"],
            "递延所得税资产": summary["total_dta"],
            "合并净利润净影响": summary["total_net_profit_impact"],
        }])
        summary_cn.to_excel(writer, sheet_name="汇总", index=False)
        rows = []
        for r in results:
            for e in r["entries"]:
                rows.append({"分录编号":r["elim_no"],"交易类型":r["trade_label"],
                             "销售方":r["seller"],"购买方":r["buyer"],"批次号":r["batch_no"],
                             "存货名称":r["item_name"],"未实现利润":r["unrealized_profit"],
                             "借方科目":e["dr"] or "","贷方科目":e["cr"] or "","金额":e["amount"],"备注":e["note"]})
        pd.DataFrame(rows).to_excel(writer, sheet_name="抵消分录明细", index=False)
    return output.getvalue()


# ========== 模板生成 ==========
def gen_equity(): return pd.DataFrame({"母公司编码":["P001","P001","P001","S001"],"子公司编码":["S001","S002","S003","S004"],"母公司持股比例":[0.80,0.60,1.00,0.70],"是否纳入合并":["是","是","是","是"]})
def gen_sales():  return pd.DataFrame({"销售方编码":["P001","S001","S001","P001"],"购买方编码":["S001","P001","S002","S003"],"存货批次号":["B202605001","B202605002","B202605003","B202605004"],"存货名称":["A产品","B材料","C部件","A产品"],"销售收入":[1000000,500000,300000,800000],"销售成本":[600000,350000,180000,480000],"销售数量":[100,200,150,80],"会计期间":["202605","202605","202605","202605"]})
def gen_inv():    return pd.DataFrame({"主体编码":["S001","P001","S002","S003"],"存货批次号":["B202605001","B202605002","B202605003","B202605004"],"期末结存数量":[70,50,120,80],"内部购入总数量":[100,200,150,80]})

def to_excel_bytes(df): out=io.BytesIO(); df.to_excel(out,index=False,engine="openpyxl"); return out.getvalue()


# ========== 分录卡片渲染 ==========
def render_entry_card(r):
    """渲染一张抵消分录卡片"""
    lines = []
    for e in r["entries"]:
        if e["dr"]: lines.append(f"借：{e['dr']}    {e['amount']:,.2f}")
        if e["cr"]: lines.append(f"    贷：{e['cr']}    {e['amount']:,.2f}")
    return "\n".join(lines)


# ========== 主界面 ==========
def main():
    st.title("📊 基于AI规则引擎的内部存货交易未实现利润自动抵消")
    st.caption("DRP系统 T11 合并报表与一键出表 — 子模块 | V2.0 模块化架构 | 2025级会计专硕")
    st.caption("小组成员：尚浩然、李金阳 | 指导老师：张金昌")

    init_db()
    conn = sqlite3.connect(str(DB_PATH))

    # ---------- 侧边栏 ----------
    with st.sidebar:
        st.header("⚙️ 配置")
        tax_rate_pct = st.slider("所得税税率（%）", 0, 50, 25, 1, format="%d%%",
                                  help="默认25%，高新技术企业15%")
        tax_rate = tax_rate_pct / 100.0

        st.divider(); st.header("📋 状态")
        eq_ok = st.session_state.get("equity_df") is not None
        sl_ok = st.session_state.get("sales_df") is not None
        inv_ok = st.session_state.get("inventory_df") is not None
        st.progress((eq_ok+sl_ok+inv_ok)/3, text=f"{eq_ok+sl_ok+inv_ok}/3")
        st.caption(f"{'✅' if eq_ok else '⬜'} 股权关系  {'✅' if sl_ok else '⬜'} 销售明细  {'✅' if inv_ok else '⬜'} 存货结存")

        if eq_ok or sl_ok or inv_ok:
            if st.button("🔄 清除全部", use_container_width=True):
                for k in ["equity_df","sales_df","inventory_df","results","summary"]:
                    st.session_state.pop(k, None)
                st.rerun()

        st.divider(); st.header("📋 模板")
        st.download_button("股权关系表", to_excel_bytes(gen_equity()), "模板-股权关系配置表.xlsx")
        st.download_button("销售明细表", to_excel_bytes(gen_sales()), "模板-内部销售明细表.xlsx")
        st.download_button("存货结存表", to_excel_bytes(gen_inv()), "模板-存货期末结存表.xlsx")

        st.divider(); st.header("🚀 演示")
        if st.button("一键加载示例", use_container_width=True):
            st.session_state["equity_df"] = gen_equity()
            st.session_state["sales_df"] = gen_sales()
            st.session_state["inventory_df"] = gen_inv()
            save_equity_to_db(conn, st.session_state["equity_df"])
            st.rerun()

        st.divider()
        st.caption("小组成员：尚浩然、李金阳")
        st.caption("指导老师：张金昌")
        st.caption("中国社会科学院大学 MPAcc 二班")

    # ---------- 初始化 ----------
    for k in ["equity_df","sales_df","inventory_df"]:
        if k not in st.session_state: st.session_state[k] = None

    # ---------- 三步上传 ----------
    st.subheader("📤 数据上传")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**① 股权关系配置表**")
        st.caption("母公司编码 | 子公司编码 | 持股比例 | 是否纳入合并")
        ef = st.file_uploader("上传", type=["xlsx","xls"], key="eq_up", label_visibility="visible")
        if ef:
            try:
                df = pd.read_excel(ef); errs = validate_equity_df(df)
                if errs: [st.error(e) for e in errs]
                else: st.session_state["equity_df"] = df; save_equity_to_db(conn, df)
            except Exception as ex: st.error(f"读取失败：{ex}")
        if st.session_state["equity_df"] is None:
            db_eq = load_equity_from_db(conn)
            if len(db_eq) > 0:
                st.session_state["equity_df"] = db_eq
                st.info(f"📌 已从历史记录自动加载 {len(db_eq)} 组股权关系（非手动上传）")
        if st.session_state["equity_df"] is not None:
            st.success(f"✅ 股权关系：{len(st.session_state['equity_df'])}组")
    with c2:
        st.markdown("**② 内部销售明细表**")
        st.caption("销售方 | 购买方 | 批次号 | 品名 | 收入 | 成本 | 数量 | 期间")
        sf = st.file_uploader("上传", type=["xlsx","xls"], key="sl_up", label_visibility="visible")
        if sf:
            try:
                df = pd.read_excel(sf); errs = validate_sales_df(df)
                if errs: [st.error(e) for e in errs]
                else: st.session_state["sales_df"] = df
            except Exception as ex: st.error(f"读取失败：{ex}")
        if st.session_state["sales_df"] is not None:
            st.success(f"✅ 销售明细：{len(st.session_state['sales_df'])}笔")
    with c3:
        st.markdown("**③ 存货期末结存表**")
        st.caption("主体编码 | 批次号 | 期末数量 | 购入总量")
        invf = st.file_uploader("上传", type=["xlsx","xls"], key="inv_up", label_visibility="visible")
        if invf:
            try:
                df = pd.read_excel(invf); errs = validate_inventory_df(df)
                if errs: [st.error(e) for e in errs]
                else: st.session_state["inventory_df"] = df
            except Exception as ex: st.error(f"读取失败：{ex}")
        if st.session_state["inventory_df"] is not None:
            st.success(f"✅ 存货结存：{len(st.session_state['inventory_df'])}条")

    # ---------- 手动录入 ----------
    with st.expander("✏️ 手动录入交易（可选）"):
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            ms = st.text_input("销售方", "P001"); mb = st.text_input("购买方", "S001")
            mbn = st.text_input("批次号", "B999")
        with mc2:
            mi = st.text_input("存货名称", "测试产品")
            mrv = st.number_input("销售收入", 0, 99999999, 1000000, 1000)
            mcs = st.number_input("销售成本", 0, 99999999, 600000, 1000)
        with mc3:
            mq = st.number_input("数量", 1, 99999, 100)
            me = st.number_input("期末结存", 0, 99999, 70)
            mp = st.text_input("期间", "202605")
        if st.button("➕ 添加到列表", use_container_width=True):
            nt = pd.DataFrame([{"销售方编码":ms,"购买方编码":mb,"存货批次号":mbn,"存货名称":mi,"销售收入":mrv,"销售成本":mcs,"销售数量":mq,"会计期间":mp}])
            ni = pd.DataFrame([{"主体编码":mb,"存货批次号":mbn,"期末结存数量":me,"内部购入总数量":mq}])
            st.session_state["sales_df"] = pd.concat([st.session_state["sales_df"],nt],ignore_index=True) if st.session_state["sales_df"] is not None else nt
            st.session_state["inventory_df"] = pd.concat([st.session_state["inventory_df"],ni],ignore_index=True) if st.session_state["inventory_df"] is not None else ni
            st.session_state.pop("results",None); st.success(f"已添加：{ms}→{mb} {mi}"); st.rerun()

    # ---------- 计算 ----------
    equity_df = st.session_state["equity_df"]
    sales_df = st.session_state["sales_df"]
    inventory_df = st.session_state["inventory_df"]

    st.divider()
    if equity_df is None or sales_df is None or inventory_df is None:
        st.warning("⚠️ 请上传三份表格（或点左侧「一键加载示例」）")
        conn.close(); return

    if st.button("🚀 执行抵消计算", type="primary", use_container_width=True):
        with st.spinner("计算中..."):
            summary, results, logs, batch_id, anomalies = run_engine(sales_df, inventory_df, equity_df, tax_rate)
            save_logs_to_db(conn, logs)
        st.session_state["summary"] = summary
        st.session_state["results"] = results
        st.session_state["anomalies"] = anomalies
        st.session_state["period"] = str(sales_df["会计期间"].iloc[0]) if len(sales_df)>0 else datetime.now().strftime("%Y%m")

    if "results" not in st.session_state: conn.close(); return

    summary = st.session_state["summary"]
    results = st.session_state["results"]
    anomalies = st.session_state.get("anomalies", [])
    period = st.session_state["period"]

    # ========== 仪表盘 ==========
    st.divider()
    st.subheader("📊 仪表盘")

    dc1, dc2, dc3, dc4, dc5 = st.columns(5)
    with dc1: st.metric("本期内部交易额", f"{sum(r['revenue'] for r in results):,.0f} 元")
    with dc2: st.metric("未实现利润总额", f"{sum(r['unrealized_profit'] for r in results):,.0f} 元")
    with dc3: st.metric("已自动抵消", f"{summary['eliminated']} 笔")
    with dc4: st.metric("冲减存货", f"{summary['total_inventory_offset']:,.0f} 元")
    with dc5: st.metric("合并净利润影响", f"{summary['total_net_profit_impact']:,.0f} 元", delta_color="inverse")

    # ========== 快捷操作 ==========
    st.divider()
    st.subheader("🚀 快捷操作")
    qa1, qa2, qa3, qa4 = st.columns(4)
    with qa1:
        st.markdown(f"""
        <div style="border:1px solid #1A3C6E;border-radius:8px;padding:12px;text-align:center;background:#f0f5ff">
            <div style="font-size:24px">📊</div>
            <div style="font-weight:600;font-size:14px;color:#1A3C6E">抵消笔数</div>
            <div style="font-size:22px;font-weight:700;color:#1A3C6E">{summary['eliminated']}</div>
            <div style="font-size:11px;color:#666">共{summary['total_trades']}笔交易</div>
        </div>""", unsafe_allow_html=True)
    with qa2:
        st.markdown(f"""
        <div style="border:1px solid #E74C3C;border-radius:8px;padding:12px;text-align:center;background:#fff5f5">
            <div style="font-size:24px">⚠️</div>
            <div style="font-weight:600;font-size:14px;color:#E74C3C">异常预警</div>
            <div style="font-size:22px;font-weight:700;color:#E74C3C">{summary.get('anomaly_count', 0)}</div>
            <div style="font-size:11px;color:#666">项数据异常</div>
        </div>""", unsafe_allow_html=True)
    with qa3:
        st.markdown(f"""
        <div style="border:1px solid #2ECC71;border-radius:8px;padding:12px;text-align:center;background:#f5fff5">
            <div style="font-size:24px">💰</div>
            <div style="font-weight:600;font-size:14px;color:#2ECC71">冲减存货</div>
            <div style="font-size:22px;font-weight:700;color:#2ECC71">¥{summary['total_inventory_offset']:,.0f}</div>
            <div style="font-size:11px;color:#666">未实现利润总额</div>
        </div>""", unsafe_allow_html=True)
    with qa4:
        st.markdown(f"""
        <div style="border:1px solid #F39C12;border-radius:8px;padding:12px;text-align:center;background:#fffdf5">
            <div style="font-size:24px">📉</div>
            <div style="font-weight:600;font-size:14px;color:#F39C12">净利润影响</div>
            <div style="font-size:22px;font-weight:700;color:#F39C12">¥{summary['total_net_profit_impact']:,.0f}</div>
            <div style="font-size:11px;color:#666">合并层面</div>
        </div>""", unsafe_allow_html=True)

    # ========== 最近操作 ==========
    st.divider()
    st.subheader("🕐 最近操作")
    try:
        recent = conn.execute(
            "SELECT elim_no, trade_json, created_at FROM calc_log ORDER BY created_at DESC LIMIT 8"
        ).fetchall()
        if recent:
            for row in recent:
                try:
                    tj = json.loads(row[1])
                    seller = tj.get('seller','?'); buyer = tj.get('buyer','?')
                    label = f"{seller}→{buyer}"
                except:
                    label = row[0]
                st.caption(f"🕐 {row[2][:16]}  **{row[0]}**  {label}")
        else:
            st.caption("暂无历史记录")
    except:
        st.caption("暂无历史记录")

    # ========== 异常预警 ==========
    if anomalies:
        st.divider()
        st.subheader(f"⚠️ 异常/警告（{len(anomalies)}项）")
        for a in anomalies:
            st.warning(a)

    # ========== 图表 ==========
    st.divider()
    ch1, ch2 = st.columns(2)
    with ch1:
        st.subheader("📊 抵消金额构成")
        cd = pd.DataFrame({"科目":["冲减存货","递延所得税资产","少数股东权益"],
                           "金额":[summary["total_inventory_offset"],summary["total_dta"],
                                   summary.get("total_minority_impact",0)]})
        cd = cd[cd["金额"]>0]
        if len(cd)>0:
            fig, ax = plt.subplots(figsize=(5, 3.5))
            colors = ['#ff6b6b', '#4ecdc4', '#ffe66d']
            ax.barh(cd["科目"], cd["金额"], color=colors[:len(cd)], height=0.5)
            ax.set_xlabel("金额（元）"); ax.invert_yaxis()
            for i, v in enumerate(cd["金额"]):
                ax.text(v + max(cd["金额"])*0.01, i, f'{v:,.0f}', va='center', fontsize=10)
            ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            fig.patch.set_facecolor('none'); ax.set_facecolor('none')
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)
        else: st.caption("无数据")
    with ch2:
        st.subheader("📊 交易类型分布")
        tc = {}
        for r in results:
            tc[r["trade_label"]] = tc.get(r["trade_label"], 0) + r["unrealized_profit"]
        if tc:
            fig, ax = plt.subplots(figsize=(5, 3.5))
            labels = list(tc.keys()); values = list(tc.values())
            colors2 = ['#00d2ff', '#7b68ee', '#ff6b6b']
            ax.barh(labels, values, color=colors2[:len(labels)], height=0.4)
            ax.set_xlabel("未实现利润（元）"); ax.invert_yaxis()
            for i, v in enumerate(values):
                ax.text(v + max(values)*0.01, i, f'{v:,.0f}', va='center', fontsize=10)
            ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            fig.patch.set_facecolor('none'); ax.set_facecolor('none')
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

    # ========== 抵消前后对比 ==========
    st.divider()
    st.subheader("📈 合并报表关键项目抵消前后对比")
    tr = sum(r["revenue"] for r in results)
    tc2 = sum(r["cost"] for r in results)
    tib = sum(r["revenue"]*r["unsold_ratio"] for r in results)
    cdata = pd.DataFrame([
        {"项目":"营业收入","抵消前":f"{tr:,.0f}","抵消金额":f"-{summary['total_revenue_offset']:,.0f}","抵消后":f"{tr-summary['total_revenue_offset']:,.0f}"},
        {"项目":"营业成本","抵消前":f"{tc2:,.0f}","抵消金额":f"-{summary['total_cost_offset']:,.0f}","抵消后":f"{tc2-summary['total_cost_offset']:,.0f}"},
        {"项目":"存货","抵消前":f"{tib:,.0f}","抵消金额":f"-{summary['total_inventory_offset']:,.0f}","抵消后":f"{tib-summary['total_inventory_offset']:,.0f}"},
        {"项目":"递延所得税资产","抵消前":"0","抵消金额":f"+{summary['total_dta']:,.0f}","抵消后":f"{summary['total_dta']:,.0f}"},
        {"项目":"净利润","抵消前":f"{tr-tc2:,.0f}","抵消金额":f"{summary['total_net_profit_impact']:,.0f}","抵消后":f"{(tr-tc2)+summary['total_net_profit_impact']:,.0f}"},
    ])
    st.dataframe(cdata, use_container_width=True, hide_index=True)

    # ========== 试算平衡表 ==========
    st.divider()
    st.subheader("⚖️ 合并试算平衡表")

    total_dr_bal = 0.0
    total_cr_bal = 0.0
    for r in results:
        for e in r["entries"]:
            if e["dr"]:
                total_dr_bal += abs(e["amount"])
            if e["cr"]:
                total_cr_bal += abs(e["amount"])

    diff_bal = round(total_dr_bal - total_cr_bal, 2)
    balanced = abs(diff_bal) < 0.01

    bc1, bc2, bc3, bc4 = st.columns(4)
    with bc1:
        st.metric("分录笔数", f"{sum(len(r['entries']) for r in results)} 笔")
    with bc2:
        st.metric("借方合计", f"¥{total_dr_bal:,.2f}")
    with bc3:
        st.metric("贷方合计", f"¥{total_cr_bal:,.2f}")
    with bc4:
        st.metric("差额", f"¥{diff_bal:,.2f}",
                  delta="✅ 借贷平衡" if balanced else "⚠️ 借贷不平",
                  delta_color="normal" if balanced else "inverse")

    # ========== 分录筛选 + 明细 ==========
    st.divider()
    st.subheader("📝 抵消分录明细")

    # -- 筛选标签 --
    trade_types = list(set(r['trade_label'] for r in results))
    ft_key = "filter_trade_type"
    if ft_key not in st.session_state:
        st.session_state[ft_key] = "全部"
    tags = ["全部"] + sorted(trade_types)
    cols = st.columns(len(tags))
    for ci, tag in enumerate(tags):
        with cols[ci]:
            btn_type = "primary" if st.session_state[ft_key] == tag else "secondary"
            if st.button(tag, key=f"ft_{tag}", type=btn_type, use_container_width=True):
                st.session_state[ft_key] = tag
                st.rerun()

    filtered = results if st.session_state[ft_key] == "全部" else [
        r for r in results if r['trade_label'] == st.session_state[ft_key]
    ]
    st.caption(f"显示 {len(filtered)}/{len(results)} 笔")

    # -- 分录卡片 --
    for i, r in enumerate(filtered):
        orig_idx = results.index(r)
        with st.expander(
            f"{r['elim_no']} | {r['trade_label']} | {r['seller']}→{r['buyer']} | "
            f"{r['item_name']} | 未实现利润：{r['unrealized_profit']:,.0f} 元",
            expanded=(i==0)
        ):
            st.caption(
                f"毛利率={r['gross_margin_rate']:.2%} | 未售={r['unsold_ratio']:.2%} | "
                f"收入={r['revenue']:,.0f} | 成本={r['cost']:,.0f}"
                + (f" | 少数股东={r['minority_ratio']:.2%}" if r['minority_ratio'] else "")
            )
            st.code(render_entry_card(r), language=None)

            # -- 手动调整分录 --
            with st.expander("✏️ 手动调整分录"):
                edited = False
                for ei, e in enumerate(r["entries"]):
                    c1, c2, c3 = st.columns([2, 2, 1])
                    with c1:
                        st.caption(f"{'借' if e['dr'] else '贷'}: {e['dr'] or e['cr']}")
                    with c2:
                        new_amt = st.number_input(
                            "金额", value=float(e["amount"]), step=1000.0,
                            key=f"edit_{orig_idx}_{ei}", format="%.2f"
                        )
                        if abs(new_amt - float(e["amount"])) > 0.001:
                            r["entries"][ei]["amount"] = new_amt
                            edited = True
                    with c3:
                        if st.button("🗑️", key=f"del_{orig_idx}_{ei}", help="删除此行"):
                            r["entries"].pop(ei)
                            st.rerun()
                if edited:
                    st.caption("✅ 金额已更新（需重新导出）")

                # 追加分录
                st.divider()
                st.caption("➕ 追加一条分录")
                ac1, ac2, ac3 = st.columns([3, 2, 2])
                with ac1:
                    new_dr = st.text_input("借方科目", key=f"ndr_{orig_idx}", placeholder="如：营业收入")
                    new_cr = st.text_input("贷方科目", key=f"ncr_{orig_idx}", placeholder="如：存货")
                with ac2:
                    new_amt2 = st.number_input("金额", value=0.0, step=1000.0, key=f"namt_{orig_idx}", format="%.2f")
                with ac3:
                    if st.button("添加", key=f"add_{orig_idx}") and new_amt2 != 0:
                        r["entries"].append({
                            "dr": new_dr or None, "cr": new_cr or None,
                            "amount": new_amt2, "note": "✏️ 人工追加"
                        })
                        st.rerun()

            # -- 删除整笔 --
            if st.button("🗑️ 删除此笔抵消", key=f"delrec_{orig_idx}"):
                st.session_state["results"].remove(r)
                st.rerun()

    # ========== 历史记录 ==========
    st.divider()
    st.subheader("📜 历史处理记录")
    try:
        batches = conn.execute(
            "SELECT batch_id, COUNT(*) as cnt, MIN(created_at) as created "
            "FROM calc_log GROUP BY batch_id ORDER BY created DESC LIMIT 10"
        ).fetchall()
        if batches:
            for b_id, cnt, created in batches:
                with st.expander(f"批次 {b_id} | {created[:16]} | {cnt} 笔分录"):
                    rows = conn.execute(
                        "SELECT elim_no, trade_json, calc_json, entries_json, log_hash, created_at "
                        "FROM calc_log WHERE batch_id = ? ORDER BY id", (b_id,)
                    ).fetchall()
                    for row in rows:
                        try:
                            tj = json.loads(row[1])
                            cj = json.loads(row[2])
                            ej = json.loads(row[3])
                            st.markdown(f"**{row[0]}** | {tj.get('seller','?')}→{tj.get('buyer','?')} | 未实现利润 ¥{cj.get('unrealized_profit',0):,.0f}")
                            lines = []
                            for e in ej:
                                if e.get('dr'): lines.append(f"借：{e['dr']}    {e['amount']:,.2f}")
                                if e.get('cr'): lines.append(f"    贷：{e['cr']}    {e['amount']:,.2f}")
                            st.code('\n'.join(lines), language=None)
                            st.caption(f"哈希：{row[4][:32]}... | {row[5][:16]}")
                        except:
                            st.caption(f"{row[0]} — 数据解析失败")
        else:
            st.caption("暂无历史记录（执行计算后自动保存）")
    except:
        st.caption("暂无历史记录")

    # ========== 导出 ==========
    st.divider()
    ec1, ec2 = st.columns(2)
    with ec1:
        st.download_button("📥 导出Excel", export_excel(results, summary),
                           f"抵消分录明细_{period}.xlsx", use_container_width=True)
    with ec2:
        st.download_button("📥 导出PDF报告", export_pdf(results, summary, period, anomalies),
                           f"抵消分录报告_{period}.pdf", use_container_width=True)

    conn.close()


if __name__ == "__main__":
    main()
