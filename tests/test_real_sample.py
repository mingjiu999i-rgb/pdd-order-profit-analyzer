from __future__ import annotations

import sqlite3
import io
import os
import tempfile
import unittest
from pathlib import Path

from pdd_analyzer.db import connect
from pdd_analyzer.cost_importer import parse_cost_file, save_cost_items
from pdd_analyzer.engine import combined_summary, daily_summary, day_summary
from pdd_analyzer.importer import Importer, ShopRequiredError, classify_fund, shop_from_filename

REAL_ORDER = os.environ.get("PDD_REAL_ORDER")
REAL_REFUND = os.environ.get("PDD_REAL_REFUND")
REAL_FUNDS_DIR = os.environ.get("PDD_REAL_FUNDS_DIR")


class RealSampleTest(unittest.TestCase):
    def test_marketing_report_header_is_supported(self):
        headers = {"日期", "成交营销花费(元)", "总营销花费(元)", "推广总花费(元)"}
        date_col, amount_col = Importer._promotion_columns(headers)
        self.assertEqual(date_col, "日期")
        self.assertEqual(amount_col, "总营销花费(元)")

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.conn = connect(Path(self.tmp.name) / "test.db")

    def tearDown(self):
        self.conn.close(); self.tmp.cleanup()

    def test_august_first_acceptance(self):
        if not all((REAL_ORDER, REAL_REFUND, REAL_FUNDS_DIR)): self.skipTest("未配置真实样本")
        imp = Importer(self.conn)
        imp.import_path(Path(REAL_ORDER), "验收店铺")
        imp.import_path(Path(REAL_REFUND), "验收店铺")
        for path in sorted(Path(REAL_FUNDS_DIR).glob("*.csv")): imp.import_path(path, "验收店铺")
        day = day_summary(self.conn, "2026-08-01")
        expected = {"order_count":306,"original_sales":4314.44,"unshipped_refund_orders":44,"unshipped_original":611.39,"effective_sales":3703.05,"shipped_refund_orders":41,"successful_aftersales":87,"shipped_refund":147.20,"other_deductions":28.25,"current_net":3527.60,"collection_gap":175.45,"eligible_orders":262,"settled_orders":262,"completion_rate":1.0}
        for key, value in expected.items(): self.assertEqual(day[key], value, key)
        self.assertAlmostEqual(day["net_settlement_rate"], .952620, places=6)
        self.assertEqual(day["settlement_difference"], 0.0)
        total = combined_summary(daily_summary(self.conn))
        self.assertEqual(total["order_count"], sum(d["order_count"] for d in daily_summary(self.conn)))
        self.assertAlmostEqual(total["net_settlement_rate"], total["current_net"] / total["effective_sales"], places=6)

    def test_overlapping_fund_import_keeps_real_repeats(self):
        if not REAL_FUNDS_DIR: self.skipTest("未配置真实样本")
        path = sorted(Path(REAL_FUNDS_DIR).glob("*.csv"))[0]
        imp = Importer(self.conn); first = imp.import_path(path, "测试店铺")[0]; second = imp.import_path(path, "测试店铺")[0]
        count = self.conn.execute("SELECT COUNT(*) n FROM fund_events").fetchone()["n"]
        self.assertEqual(count, first["new"])
        self.assertEqual(second["duplicate"], first["new"])
        self.assertGreater(self.conn.execute("SELECT COUNT(*) n FROM raw_rows").fetchone()["n"], count)

    def test_classification(self):
        self.assertEqual(classify_fund("退款", "0020002|交易退款-订单退款"), "refund")
        self.assertEqual(classify_fund("其他", "推广账户充值"), "transfer")
        self.assertEqual(classify_fund("扣款", "推广费扣款"), "promotion_spend")
        self.assertEqual(classify_fund("其他服务", "0050002|服务支出-消费者体验提升计划"), "other_operating")
        self.assertEqual(classify_fund("其他", "从未见过的项目"), "unknown")
        self.assertEqual(shop_from_filename("示例店铺8月份订单.csv"), "示例店铺")
        self.assertEqual(shop_from_filename("示例店铺截止9月11日退款.xlsx"), "示例店铺")
        self.assertIsNone(shop_from_filename("pdd-mall-bill-detail(xxx).csv"))
        self.assertIsNone(shop_from_filename("cb9956590c8940550aa9adf8461e4ab7orders_export2026-09-15.csv"))
        self.assertIsNone(shop_from_filename("cb9956590c8940550aa9adf8461e4ab7_20260915222321.xlsx"))

    def test_shop_detection_never_guesses_the_only_existing_shop(self):
        self.conn.execute("INSERT INTO orders(order_id,shop_name,merchant_receipt_cents,source_batch_id,updated_at) VALUES('old-order','老店铺',0,1,'now')")
        with self.assertRaises(ShopRequiredError):
            Importer(self.conn)._detect_shop("pdd-mall-bill-detail.csv", "fund", [{"商户订单号": "new-order"}])

    def test_sku_costs_are_isolated_by_shop(self):
        for order_id, shop in (("a", "甲店"), ("b", "乙店")):
            self.conn.execute(
                """INSERT INTO orders(order_id,pay_time,pay_date,shop_name,merchant_receipt_cents,
                   order_status,sku_code,quantity,source_batch_id,updated_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (order_id, "2026-09-01 10:00:00", "2026-09-01", shop, 1000,
                 "已发货", "SAME-SKU", 2, 1, "now"),
            )
        save_cost_items(self.conn, "甲店", [{"sku_key": "SAME-SKU", "unit_cost_cents": 100}])
        save_cost_items(self.conn, "乙店", [{"sku_key": "SAME-SKU", "unit_cost_cents": 300}])
        self.conn.commit()
        self.assertEqual(day_summary(self.conn, "2026-09-01", "甲店")["product_cost"], 2.0)
        self.assertEqual(day_summary(self.conn, "2026-09-01", "乙店")["product_cost"], 6.0)

    def test_cost_file_requires_sku_and_total_cost(self):
        from openpyxl import Workbook

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["SKU编码", "总成本", "快递费"])
        sheet.append(["001-A", 12.345, 2])
        stream = io.BytesIO(); workbook.save(stream)
        rows = parse_cost_file("成本.xlsx", stream.getvalue())
        self.assertEqual(rows, [{"sku_key": "001-A", "unit_cost_cents": 1235}])
        with self.assertRaisesRegex(ValueError, "不支持按零散成本导入"):
            parse_cost_file("成本.csv", "SKU编码,快递费\n001-A,2\n".encode("utf-8"))

    def test_new_order_template_uses_transaction_time(self):
        data = (
            "订单号,订单状态,商家实收金额(元),商品数量(件),商品id,商品规格,"
            "商家编码-规格维度,售后状态,订单成交时间\n"
            "260915-test,已发货,19.50,2,123,原味两袋,SKU-NEW,无售后,2026-09-15 23:06:56\n"
        ).encode("utf-8")
        result = Importer(self.conn).import_bytes("orders_export.csv", data, "测试店铺")
        order = self.conn.execute("SELECT * FROM orders WHERE order_id='260915-test'").fetchone()
        self.assertEqual(result["report_type"], "order")
        self.assertEqual(order["pay_time"], "2026-09-15 23:06:56")
        self.assertEqual(order["pay_date"], "2026-09-15")
        self.assertEqual(order["sku_code"], "SKU-NEW")


if __name__ == "__main__": unittest.main()
