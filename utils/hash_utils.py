"""
哈希工具模块
提供SHA256哈希计算和哈希链功能
"""
import hashlib
import json


def compute_hash(prev_hash: str, data_str: str) -> str:
    """计算SHA256哈希（含上一笔哈希，形成哈希链）"""
    content = f"{prev_hash}|{data_str}"
    return hashlib.sha256(content.encode("utf-8")).hexdigest()[:16]


def compute_data_hash(data: dict) -> str:
    """对字典数据计算SHA256哈希"""
    data_str = json.dumps(data, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(data_str.encode("utf-8")).hexdigest()[:16]


def verify_chain(hashes: list) -> bool:
    """验证哈希链完整性"""
    for i in range(1, len(hashes)):
        expected = compute_hash(hashes[i - 1], str(i))
        if hashes[i] != expected:
            return False
    return True
