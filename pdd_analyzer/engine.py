from __future__ import annotations

import sqlite3


def yuan(cents: int | None) -> float | None:
    return None if cents is None else round(cents / 100, 2)


def daily_summary(conn: sqlite3.Connection, start_date: str | None = None, end_date: str | None = None, shop_name: str | None = None) -> list[dict]:
    where, params = ["pay_date IS NOT NULL"], []
    if start_date: where.append("pay_date>=?"); params.append(start_date)
    if end_date: where.append("pay_date<=?"); params.append(end_date)
    if shop_name: where.append("shop_name=?"); params.append(shop_name)
    dates = conn.execute(f"SELECT DISTINCT shop_name,pay_date FROM orders WHERE {' AND '.join(where)} ORDER BY pay_date DESC,shop_name", params).fetchall()
    return [day_summary(conn, r["pay_date"], r["shop_name"]) for r in dates]


def combined_summary(days: list[dict]) -> dict | None:
    if not days:
        return None
    amount_fields = ["original_sales", "effective_sales", "shipped_refund", "other_deductions", "current_net"]
    count_fields = ["order_count", "eligible_orders", "settled_orders"]
    total = {key: round(sum((d.get(key) or 0) for d in days), 2) for key in amount_fields}
    total.update({key: sum(int(d.get(key) or 0) for d in days) for key in count_fields})
    shops = sorted({d.get("shop_name") or "未识别店铺" for d in days})
    dates = sorted(d["date"] for d in days)
    total.update({
        "shop_name": shops[0] if len(shops) == 1 else "全部店铺",
        "date": f"{dates[0]} 至 {dates[-1]}",
        "completion_rate": round(total["settled_orders"] / total["eligible_orders"], 6) if total["eligible_orders"] else 1.0,
        "net_settlement_rate": round(total["current_net"] / total["effective_sales"], 6) if total["effective_sales"] else None,
        "mature": all(d["mature"] for d in days),
        "unknown_funds": sum(d.get("unknown_funds") or 0 for d in days),
        "settlement_difference": None,
    })
    return total


def combined_summaries(days: list[dict]) -> list[dict]:
    shops = sorted({d.get("shop_name") or "未识别店铺" for d in days})
    return [combined_summary([d for d in days if (d.get("shop_name") or "未识别店铺") == shop]) for shop in shops]


def day_summary(conn: sqlite3.Connection, pay_date: str, shop_name: str | None = None) -> dict:
    if shop_name is None:
        orders = conn.execute("SELECT * FROM orders WHERE pay_date=?", (pay_date,)).fetchall()
    else:
        orders = conn.execute("SELECT * FROM orders WHERE pay_date=? AND shop_name=?", (pay_date, shop_name)).fetchall()
    ids = [o["order_id"] for o in orders]
    if not ids:
        return {"date": pay_date, "order_count": 0}
    placeholders = ",".join("?" * len(ids))
    refunds = conn.execute(f"SELECT * FROM refunds WHERE order_id IN ({placeholders})", ids).fetchall()
    funds = conn.execute(f"SELECT * FROM fund_events WHERE order_id IN ({placeholders})", ids).fetchall()
    successful = [r for r in refunds if r["aftersale_status"] == "退款成功"]
    unshipped_ids = {r["order_id"] for r in successful if r["shipping_stage"] == "未发货"}
    shipped_ids = {r["order_id"] for r in successful if r["shipping_stage"] == "已发货"}
    original = sum(o["merchant_receipt_cents"] for o in orders)
    unshipped_original = sum(o["merchant_receipt_cents"] for o in orders if o["order_id"] in unshipped_ids)
    effective = original - unshipped_original
    positive_ids = {f["order_id"] for f in funds if f["category"] == "positive_settlement" and f["income_cents"] > 0}
    eligible_ids = {o["order_id"] for o in orders if o["order_id"] not in unshipped_ids and "取消" not in (o["order_status"] or "")}
    operating = [f for f in funds if f["category"] in {"positive_settlement", "refund", "other_operating"}]
    current_net = sum(f["income_cents"] + f["expense_cents"] for f in operating)
    shipped_refund = -sum(f["income_cents"] + f["expense_cents"] for f in funds if f["category"] == "refund" and f["order_id"] in shipped_ids)
    other_net = sum(f["income_cents"] + f["expense_cents"] for f in funds if f["category"] == "other_operating")
    other_deductions = -other_net
    completion = len(eligible_ids & positive_ids) / len(eligible_ids) if eligible_ids else 1.0
    settlement_rate = current_net / effective if effective else None
    unknown = sum(1 for f in funds if f["category"] == "unknown")
    expected_net = effective - shipped_refund - other_deductions
    difference = current_net - expected_net if completion >= 0.999999 else None
    promotion_rows = [f for f in funds if f["category"] == "promotion_spend"]
    promotion = -sum(f["income_cents"] + f["expense_cents"] for f in promotion_rows) if promotion_rows else None
    return {
        "date": pay_date, "shop_name": shop_name or (orders[0]["shop_name"] if orders else "未识别店铺"), "order_count": len(orders), "original_sales": yuan(original),
        "unshipped_refund_orders": len(unshipped_ids), "unshipped_original": yuan(unshipped_original),
        "effective_sales": yuan(effective), "shipped_refund_orders": len(shipped_ids),
        "successful_aftersales": len(successful), "shipped_refund": yuan(shipped_refund),
        "other_deductions": yuan(other_deductions), "current_net": yuan(current_net),
        "promotion_fee": yuan(promotion), "eligible_orders": len(eligible_ids),
        "settled_orders": len(eligible_ids & positive_ids), "completion_rate": round(completion, 6),
        "net_settlement_rate": None if settlement_rate is None else round(settlement_rate, 6),
        "mature": completion >= 0.999999, "unknown_funds": unknown,
        "settlement_difference": yuan(difference),
    }


def day_detail(conn: sqlite3.Connection, pay_date: str, shop_name: str | None = None) -> dict:
    summary = day_summary(conn, pay_date, shop_name)
    if shop_name is None:
        orders = [dict(r) for r in conn.execute("SELECT * FROM orders WHERE pay_date=? ORDER BY pay_time,order_id", (pay_date,))]
    else:
        orders = [dict(r) for r in conn.execute("SELECT * FROM orders WHERE pay_date=? AND shop_name=? ORDER BY pay_time,order_id", (pay_date, shop_name))]
    ids = [o["order_id"] for o in orders]
    if not ids:
        return {"summary": summary, "orders": [], "refunds": [], "funds": [], "issues": []}
    marks = ",".join("?" * len(ids))
    refunds = [dict(r) for r in conn.execute(f"SELECT * FROM refunds WHERE order_id IN ({marks}) ORDER BY apply_time", ids)]
    funds = [dict(r) for r in conn.execute(f"SELECT * FROM fund_events WHERE order_id IN ({marks}) ORDER BY occurred_at,id", ids)]
    issues = []
    if summary["unknown_funds"]:
        issues.append({"type": "未知资金类型", "count": summary["unknown_funds"]})
    if summary["settlement_difference"] not in (None, 0.0):
        issues.append({"type": "结算差异", "amount": summary["settlement_difference"]})
    missing = summary["eligible_orders"] - summary["settled_orders"]
    if missing:
        issues.append({"type": "尚未入账订单", "count": missing})
    return {"summary": summary, "orders": orders, "refunds": refunds, "funds": funds, "issues": issues}


def audit_summary(conn: sqlite3.Connection, shop_name: str | None = None) -> dict:
    shop_filter = " AND r.shop_name=?" if shop_name else ""
    unmatched_refunds = conn.execute("SELECT COUNT(*) n FROM refunds r LEFT JOIN orders o ON o.order_id=r.order_id WHERE o.order_id IS NULL" + shop_filter, ((shop_name,) if shop_name else ())).fetchone()["n"]
    fund_filter = " AND f.shop_name=?" if shop_name else ""
    unmatched_funds = conn.execute("SELECT COUNT(*) n FROM fund_events f LEFT JOIN orders o ON o.order_id=f.order_id WHERE o.order_id IS NULL AND f.category NOT IN ('transfer','promotion_spend')" + fund_filter, ((shop_name,) if shop_name else ())).fetchone()["n"]
    unknown = conn.execute("SELECT COUNT(*) n FROM fund_events WHERE category='unknown'" + (" AND shop_name=?" if shop_name else ""), ((shop_name,) if shop_name else ())).fetchone()["n"]
    if shop_name:
        batches = [dict(r) for r in conn.execute("SELECT * FROM import_batches WHERE shop_name=? ORDER BY id DESC LIMIT 30", (shop_name,))]
    else:
        batches = [dict(r) for r in conn.execute("SELECT * FROM import_batches ORDER BY id DESC LIMIT 30")]
    return {"unmatched_refunds": unmatched_refunds, "unmatched_funds": unmatched_funds, "unknown_funds": unknown, "batches": batches}
