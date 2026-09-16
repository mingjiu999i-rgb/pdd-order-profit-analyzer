from __future__ import annotations

import sqlite3
from pathlib import Path


SCHEMA = """
PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS import_batches (
  id INTEGER PRIMARY KEY, file_name TEXT NOT NULL, file_sha256 TEXT NOT NULL,
  report_type TEXT NOT NULL, shop_name TEXT, imported_at TEXT NOT NULL,
  source_start TEXT, source_end TEXT, total_rows INTEGER NOT NULL DEFAULT 0,
  new_rows INTEGER NOT NULL DEFAULT 0, updated_rows INTEGER NOT NULL DEFAULT 0,
  existing_rows INTEGER NOT NULL DEFAULT 0, duplicate_rows INTEGER NOT NULL DEFAULT 0,
  suspected_rows INTEGER NOT NULL DEFAULT 0, error TEXT
);
CREATE TABLE IF NOT EXISTS raw_rows (
  id INTEGER PRIMARY KEY, batch_id INTEGER NOT NULL, report_type TEXT NOT NULL,
  row_number INTEGER NOT NULL, row_json TEXT NOT NULL, row_signature TEXT NOT NULL,
  event_ordinal INTEGER NOT NULL DEFAULT 1,
  FOREIGN KEY(batch_id) REFERENCES import_batches(id)
);
CREATE TABLE IF NOT EXISTS orders (
  order_id TEXT PRIMARY KEY, pay_time TEXT, pay_date TEXT,
  shop_name TEXT,
  merchant_receipt_cents INTEGER NOT NULL, order_status TEXT, aftersale_status TEXT,
  product_id TEXT, sku TEXT, sku_code TEXT, quantity INTEGER, source_batch_id INTEGER NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS refunds (
  aftersale_id TEXT PRIMARY KEY, order_id TEXT NOT NULL, trade_amount_cents INTEGER,
  shop_name TEXT,
  refund_amount_cents INTEGER NOT NULL, aftersale_status TEXT, shipping_stage TEXT,
  refund_type TEXT, apply_time TEXT, agree_time TEXT, source_batch_id INTEGER NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS fund_events (
  id INTEGER PRIMARY KEY, event_key TEXT UNIQUE NOT NULL, row_signature TEXT NOT NULL,
  event_ordinal INTEGER NOT NULL, order_id TEXT, shop_name TEXT, occurred_at TEXT NOT NULL,
  income_cents INTEGER NOT NULL, expense_cents INTEGER NOT NULL,
  account_type TEXT, business_code TEXT, business_description TEXT, note TEXT,
  category TEXT NOT NULL, source_batch_id INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS fund_event_sources (
  fund_event_id INTEGER NOT NULL, raw_row_id INTEGER NOT NULL UNIQUE,
  PRIMARY KEY(fund_event_id, raw_row_id)
);
CREATE TABLE IF NOT EXISTS sku_costs (
  shop_name TEXT NOT NULL, sku_key TEXT NOT NULL,
  unit_cost_cents INTEGER NOT NULL, updated_at TEXT NOT NULL,
  PRIMARY KEY(shop_name, sku_key)
);
CREATE TABLE IF NOT EXISTS promotion_expenses (
  id INTEGER PRIMARY KEY, row_key TEXT UNIQUE NOT NULL,
  shop_name TEXT NOT NULL, spend_date TEXT NOT NULL,
  amount_cents INTEGER NOT NULL, source_batch_id INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_orders_pay_date ON orders(pay_date);
CREATE INDEX IF NOT EXISTS idx_refunds_order ON refunds(order_id);
CREATE INDEX IF NOT EXISTS idx_funds_order ON fund_events(order_id);
CREATE INDEX IF NOT EXISTS idx_promotion_shop_date ON promotion_expenses(shop_name, spend_date);
"""


def connect(path: str | Path) -> sqlite3.Connection:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    # 兼容第一版已经创建的数据库。
    for table in ("import_batches", "orders", "refunds", "fund_events"):
        columns = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
        if "shop_name" not in columns:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN shop_name TEXT")
    order_columns = {r["name"] for r in conn.execute("PRAGMA table_info(orders)")}
    if "sku_code" not in order_columns:
        conn.execute("ALTER TABLE orders ADD COLUMN sku_code TEXT")
        for row in conn.execute("SELECT row_json FROM raw_rows WHERE report_type='order'"):
            try:
                import json
                source = json.loads(row["row_json"])
                order_id = str(source.get("订单号", "")).strip()
                code = str(source.get("商家编码-规格维度", "")).strip()
                if order_id and code:
                    conn.execute("UPDATE orders SET sku_code=? WHERE order_id=?", (code, order_id))
            except (ValueError, TypeError):
                continue
    cost_columns = {r["name"] for r in conn.execute("PRAGMA table_info(sku_costs)")}
    if "shop_name" not in cost_columns:
        # 第一版成本表只有 SKU 一个维度。升级时把已有成本复制到实际使用该
        # SKU 的各店铺，之后每个店铺可以维护自己的成本，互不影响。
        conn.execute("ALTER TABLE sku_costs RENAME TO sku_costs_legacy")
        conn.execute("""
            CREATE TABLE sku_costs (
              shop_name TEXT NOT NULL, sku_key TEXT NOT NULL,
              unit_cost_cents INTEGER NOT NULL, updated_at TEXT NOT NULL,
              PRIMARY KEY(shop_name, sku_key)
            )
        """)
        conn.execute("""
            INSERT OR IGNORE INTO sku_costs(shop_name,sku_key,unit_cost_cents,updated_at)
            SELECT DISTINCT o.shop_name,c.sku_key,c.unit_cost_cents,c.updated_at
            FROM sku_costs_legacy c JOIN orders o ON TRIM(o.sku_code)=c.sku_key
            WHERE TRIM(COALESCE(o.shop_name,''))<>''
        """)
    conn.commit()
    return conn
