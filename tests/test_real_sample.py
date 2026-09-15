from __future__ import annotations

import sqlite3
import os
import tempfile
import unittest
from pathlib import Path

from pdd_analyzer.db import connect
from pdd_analyzer.engine import combined_summary, daily_summary, day_summary
from pdd_analyzer.importer import Importer, ShopRequiredError, classify_fund, shop_from_filename

REAL_ORDER = os.environ.get("PDD_REAL_ORDER")
REAL_REFUND = os.environ.get("PDD_REAL_REFUND")
REAL_FUNDS_DIR = os.environ.get("PDD_REAL_FUNDS_DIR")


class RealSampleTest(unittest.TestCase):
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
        expected = {"order_count":306,"original_sales":4314.44,"unshipped_refund_orders":44,"unshipped_original":611.39,"effective_sales":3703.05,"shipped_refund_orders":41,"successful_aftersales":87,"shipped_refund":147.20,"other_deductions":28.25,"current_net":3527.60,"eligible_orders":262,"settled_orders":262,"completion_rate":1.0}
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

    def test_shop_detection_never_guesses_the_only_existing_shop(self):
        self.conn.execute("INSERT INTO orders(order_id,shop_name,merchant_receipt_cents,source_batch_id,updated_at) VALUES('old-order','老店铺',0,1,'now')")
        with self.assertRaises(ShopRequiredError):
            Importer(self.conn)._detect_shop("pdd-mall-bill-detail.csv", "fund", [{"商户订单号": "new-order"}])


if __name__ == "__main__": unittest.main()
