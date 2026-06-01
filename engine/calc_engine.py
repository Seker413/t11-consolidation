"""
未实现利润计算引擎
"""


def calculate_unrealized_profit(revenue: float, cost: float,
                                quantity: int, ending_qty: int) -> tuple:
    """
    计算未实现利润
    返回: (gross_margin_rate, unsold_ratio, unrealized_profit)
    """
    if revenue <= 0 or quantity <= 0:
        return 0, 0, 0

    gross_margin_rate = round((revenue - cost) / revenue, 6)
    unsold_ratio = round(ending_qty / quantity, 6) if quantity > 0 else 0
    unrealized_profit = round(revenue * gross_margin_rate * unsold_ratio, 2)

    return gross_margin_rate, unsold_ratio, unrealized_profit


def safe_eval(formula: str, context: dict) -> float:
    """安全计算公式（受限命名空间，防止代码注入）"""
    allowed = {
        "revenue": context.get("revenue", 0),
        "cost": context.get("cost", 0),
        "unrealized_profit": context.get("unrealized_profit", 0),
        "unsold_ratio": context.get("unsold_ratio", 0),
        "tax_rate": context.get("tax_rate", 0.25),
        "minority_ratio": context.get("minority_ratio", 0),
        "round": round,
        "abs": abs,
        "min": min,
        "max": max,
    }
    return eval(formula, {"__builtins__": {}}, allowed)


def compute_summary(results: list, tax_rate: float, sales_count: int,
                    anomaly_count: int = 0, skipped_count: int = 0) -> dict:
    """从结果列表汇总统计数据"""
    summary = {
        "total_trades": sales_count,
        "eliminated": len(results),
        "skipped": skipped_count,
        "anomaly_count": anomaly_count,
        "total_revenue_offset": 0.0,
        "total_cost_offset": 0.0,
        "total_inventory_offset": 0.0,
        "total_dta": 0.0,
        "total_net_profit_impact": 0.0,
        "total_minority_impact": 0.0,
    }

    for r in results:
        for e in r["entries"]:
            if e["dr"] == "营业收入":
                summary["total_revenue_offset"] += abs(e["amount"])
            if e["cr"] == "营业成本":
                summary["total_cost_offset"] += abs(e["amount"])
            if e["cr"] == "存货":
                summary["total_inventory_offset"] += abs(e["amount"])
            if e["dr"] == "递延所得税资产":
                summary["total_dta"] += abs(e["amount"])
            if e["dr"] == "少数股东权益":
                summary["total_minority_impact"] += abs(e["amount"])

    summary["total_net_profit_impact"] = round(
        -summary["total_inventory_offset"] * (1 - tax_rate), 2
    )
    return summary
