import tempfile
import unittest
from pathlib import Path

from pdd_analyzer.db import clear_all_data, connect


class ClearDataTest(unittest.TestCase):
    def test_clears_business_tables_and_keeps_schema_usable(self):
        with tempfile.TemporaryDirectory() as folder:
            db_path = Path(folder) / "analysis.db"
            with connect(db_path) as conn:
                batch_id = conn.execute(
                    """INSERT INTO import_batches
                    (file_name,file_sha256,report_type,shop_name,imported_at,total_rows)
                    VALUES ('订单.csv','hash','order','测试店铺','2026-09-16',1)"""
                ).lastrowid
                conn.execute(
                    """INSERT INTO orders
                    (order_id,pay_time,pay_date,shop_name,merchant_receipt_cents,
                     source_batch_id,updated_at)
                    VALUES ('123','2026-09-16 10:00:00','2026-09-16','测试店铺',1000,?,
                    '2026-09-16')""",
                    (batch_id,),
                )
                conn.execute(
                    "INSERT INTO sku_costs VALUES ('测试店铺','SKU-1',300,'2026-09-16')"
                )
                conn.execute(
                    "INSERT INTO promotion_expenses(row_key,shop_name,spend_date,amount_cents,source_batch_id) VALUES ('p1','测试店铺','2026-09-16',200,?)",
                    (batch_id,),
                )
                conn.commit()

                counts = clear_all_data(conn)

                self.assertEqual(counts["orders"], 1)
                self.assertEqual(counts["import_batches"], 1)
                self.assertEqual(counts["sku_costs"], 1)
                self.assertEqual(counts["promotion_expenses"], 1)
                self.assertEqual(conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0], 0)
                self.assertEqual(conn.execute("SELECT COUNT(*) FROM import_batches").fetchone()[0], 0)
                self.assertEqual(conn.execute("SELECT COUNT(*) FROM sku_costs").fetchone()[0], 0)
                self.assertEqual(conn.execute("SELECT COUNT(*) FROM promotion_expenses").fetchone()[0], 0)
                conn.execute(
                    """INSERT INTO import_batches
                    (file_name,file_sha256,report_type,imported_at,total_rows)
                    VALUES ('重新导入.csv','new-hash','order','2026-09-16',0)"""
                )
                conn.commit()
            conn.close()


if __name__ == "__main__":
    unittest.main()
