"""
抵消分录生成器
基于JSON规则引擎，根据交易类型和计算参数自动生成合并抵消分录
"""
import json
from pathlib import Path
from engine.calc_engine import safe_eval

RULES_PATH = Path(__file__).parent.parent / "rules.json"


def load_rules() -> dict:
    """加载JSON规则配置"""
    with open(RULES_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def generate_entries(trade_type: str, context: dict,
                     rules: dict = None) -> list:
    """
    根据规则引擎生成抵消分录
    返回: [{"dr": ..., "cr": ..., "amount": ..., "note": ...}, ...]
    """
    if rules is None:
        rules = load_rules()

    templates = rules["entry_templates"]
    if trade_type not in templates:
        return []

    template = templates[trade_type]
    entries = []

    for rule in template["entries"]:
        amount = round(safe_eval(rule["amount_formula"], context), 2)
        if amount == 0:
            continue
        entries.append({
            "dr": rule.get("dr"),
            "cr": rule.get("cr"),
            "amount": amount,
            "note": rule.get("note", ""),
        })

    for rule in template.get("minority_entries", []):
        amount = round(safe_eval(rule["amount_formula"], context), 2)
        if amount == 0:
            continue
        entries.append({
            "dr": rule.get("dr"),
            "cr": rule.get("cr"),
            "amount": amount,
            "note": rule.get("note", ""),
        })

    return entries


def get_trade_label(trade_type: str) -> str:
    """获取交易类型中文标签"""
    rules = load_rules()
    return rules["entry_templates"].get(trade_type, {}).get("label", trade_type)
