"""
交易类型分类器
负责根据股权关系图谱判定内部交易类型（顺流/逆流/平流）
"""


def build_equity_graph(equity_df):
    """
    构建股权关系图谱
    返回:
      parent_of: {child_code: parent_code}
      ratio_of: {child_code: share_ratio}
      parents: set of all parent codes
      children: set of all child codes
    """
    parent_of = {}
    ratio_of = {}
    parents = set()
    children = set()

    for _, row in equity_df.iterrows():
        p = str(row["母公司编码"]).strip()
        c = str(row["子公司编码"]).strip()
        r = float(row["母公司持股比例"])
        parent_of[c] = p
        ratio_of[c] = r
        parents.add(p)
        children.add(c)

    return parent_of, ratio_of, parents, children


def classify_trade(seller: str, buyer: str, parent_of: dict, parents: set, children: set) -> tuple:
    """
    判定交易类型
    返回: (trade_type, minority_ratio)
      trade_type: DOWN/UP/FLAT/UNKNOWN
      minority_ratio: 少数股东比例（DOWN为None）
    """
    seller = str(seller).strip()
    buyer = str(buyer).strip()

    # 规则1: seller是buyer的母公司 → 顺流
    if buyer in parent_of and parent_of[buyer] == seller:
        return "DOWN", None

    # 规则2: buyer是seller的母公司 → 逆流
    if seller in parent_of and parent_of[seller] == buyer:
        return "UP", None

    # 规则3: 同一母公司 → 平流
    if seller in parent_of and buyer in parent_of:
        if parent_of[seller] == parent_of[buyer]:
            return "FLAT", None

    return "UNKNOWN", None


def get_minority_ratio(trade_type: str, seller: str, buyer: str,
                       parent_of: dict, ratio_of: dict) -> float:
    """获取少数股东比例"""
    seller = str(seller).strip()

    if trade_type == "DOWN":
        return None
    elif trade_type in ("UP", "FLAT"):
        if seller in ratio_of:
            return round(1 - ratio_of[seller], 4)
        return 0.0
    return None


def detect_anomalies(sales_df, inventory_df, equity_df, parent_of):
    """
    检测数据异常
    返回: (anomaly_messages, anomaly_count)
    """
    anomalies = []

    # 1. 销售有但库存无
    for _, trade in sales_df.iterrows():
        batch = str(trade["存货批次号"])
        buyer = str(trade["购买方编码"]).strip()
        inv_match = inventory_df[
            (inventory_df["主体编码"].astype(str).str.strip() == buyer) &
            (inventory_df["存货批次号"].astype(str).str.strip() == batch)
        ]
        if len(inv_match) == 0:
            anomalies.append(f"⚠ 批次 {batch}（{buyer}购入）在存货结存表中无记录，可能已全部售出或数据缺失")

    # 2. 毛利率异常（负毛利率）
    for _, trade in sales_df.iterrows():
        rev = float(trade["销售收入"])
        cost = float(trade["销售成本"])
        if rev > 0 and cost > rev:
            batch = str(trade["存货批次号"])
            anomalies.append(f"⚠ 批次 {batch} 毛利率为负（收入{rev:,.0f} < 成本{cost:,.0f}），可能为亏损销售或数据错误")

    # 3. 未匹配股权关系
    all_known = set(parent_of.keys()) | set(parent_of.values())
    for _, trade in sales_df.iterrows():
        s = str(trade["销售方编码"]).strip()
        b = str(trade["购买方编码"]).strip()
        if s not in all_known and b not in all_known:
            anomalies.append(f"⚠ 交易 {s}→{b} 双方均不在股权关系表中，无法判定交易类型")
        elif s not in all_known:
            anomalies.append(f"⚠ 销售方 {s} 不在股权关系表中，无法判定交易类型")
        elif b not in all_known:
            anomalies.append(f"⚠ 购买方 {b} 不在股权关系表中，无法判定交易类型")

    # 4. 期末结存 > 购入总量
    for _, inv in inventory_df.iterrows():
        end_qty = int(inv["期末结存数量"])
        total_qty = int(inv["内部购入总数量"])
        if end_qty > total_qty:
            anomalies.append(f"⚠ 批次 {inv['存货批次号']} 期末结存({end_qty}) > 购入总量({total_qty})，数据异常")

    return anomalies, len(anomalies)
