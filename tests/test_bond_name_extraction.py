import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = ROOT / "提取脚本.py"


def load_extractor():
    spec = importlib.util.spec_from_file_location("bond_extractor", SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


extractor = load_extractor()


class BondNameExtractionTests(unittest.TestCase):
    def test_name_before_code_is_kept(self):
        text = (
            "7.2交易所 26济新K2 283069.SH 1000 100.001 广发证券i020055109 "
            "to 东吴安享18M4号i020019312 约定号：169"
        )
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["债券简称"], "26济新K2")

    def test_name_after_code_beats_buy_marker(self):
        text = (
            "永宁1号i020025816 买入 283145.SH 26横琴03 4600 净价100 "
            "from 粤财信托锐益2号 i020012906 约定号457"
        )
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["债券简称"], "26横琴03")

    def test_existing_special_name_shapes_still_work(self):
        samples = {
            "245523.SH 山能YK02 1000 净价100 中信信托信昱13号 i020055109": "山能YK02",
            "268867.SH 通16优A3 3000 净价100 粤财信托添添益1号 i020012906": "通16优A3",
            "520230.SZ 26苏园GY02 2000 净价100 中信信托睿兴5号": "26苏园GY02",
        }
        for text, expected in samples.items():
            with self.subTest(expected=expected):
                rows = extractor.parse_text(text)
                self.assertEqual(rows[0]["债券简称"], expected)

    def test_all_caps_and_chinese_digit_names_are_kept(self):
        samples = {
            "245542.SH 26SACF02 1000 净价100 中信信托华盈添利4号 i020070111 出给山西证券 i020005011 约定号 856": "26SACF02",
            "244885.SH 26CHNG3K 净价100 1.6E 广发证券【Z06308】to 长城证券Z13311 约定号222": "26CHNG3K",
            "268215.SH 金玉优01 4000 净价100 粤财信托添添益 1号 Z08705 出给 华润信托瑞合5号 中信Z06308 约定号 907": "金玉优01",
        }
        for text, expected in samples.items():
            with self.subTest(expected=expected):
                rows = extractor.parse_text(text)
                self.assertEqual(rows[0]["债券简称"], expected)

    def test_mine_account_fuzzy_forms_are_normalized(self):
        samples = {
            "要素已定\n信昱13\n244885.SH 26CHNG3K 净价100 1.6E 广发证券【Z06308】to 长城证券Z13311\n约定号222": "中信信托信昱13号",
            "268215.SH 金玉优01 4000 净价100 粤财信托添添益 1号 Z08705 出给 华润信托瑞合5号 中信Z06308 约定号 907": "粤财信托添添益1号",
            "268215.SH 金玉优01 7000 净价100 粤财信托添盈 1号 Z08705 出给 华润信托瑞合10号 中信Z06308 约定号 909": "粤财信托添盈1号",
        }
        for text, expected in samples.items():
            with self.subTest(expected=expected):
                rows = extractor.parse_text(text)
                self.assertEqual(rows[0]["我方账户"], expected)

    def test_multi_rows_header_yd_is_not_misread_as_real_yd(self):
        text = """①买入 4.96Y 520260.SZ 26文体02 100净价 100净价 2000W 07.02交易所 广发证券 to 世纪证券资管 i020000604
买入方
交易商代码：000039 世纪证券
交易员号: 00130016 李南阳
交易主体代码/简称/量/约定号：
3600062601 鑫享世成24M013号 500w 约定号12211221
3600063165 鑫享世成24M014号 500w 约定号13311331
3600051079 兴瑞世成16号 250w 约定号14411441
3600064045 兴瑞世成55号 750w 约定号15511551
卖出方
中信信托信昱11号
交易商号: 000262
交易员号：007A0001
交易主体代码：3600000001
交易主体名称：中信证券股份有限公司机构经纪
卖家先发"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 4)
        self.assertTrue(all(not row["备注"] for row in rows))
        self.assertEqual([row["约定号"] for row in rows], ["12211221", "13311331", "14411441", "15511551"])

    def test_plus_zero_date_is_not_treated_as_size_split(self):
        text = """2026/7/1 +0 520274.SZ 26豫峡K3 净价100 3000w 广发 to 首创
首创证券
交易商代码：000613
交易主体代码： 3600064456 创赢JS01号 约定号 11111111
交易主体代码： 3600064148 创赢XY01号 约定号 22222222
交易主体代码： 3600011496 创赢增利2号 约定号 33333333
交易员号：00H10004 张跃曦
中信信托信昱11号 交易商号: 000262 交易员号：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪 卖出 卖出先发"""
        yds, sizes = extractor.expand_splits(text)
        self.assertEqual(yds, ["11111111", "22222222", "33333333"])
        self.assertEqual(sizes, [])

    def test_shared_prefix_rows_can_identify_mine_price_and_size(self):
        text = """0414 云夏3号 Z07119 卖出 2.564% 99.814
281788.SH 26溧供Y1 买入 1000 中粮佳盈1号 Z08705 约定号 788

281788.SH 26溧供Y1 买入 1000 广粤尊享77号 Z08705 约定号 789"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["我方账户"], "中粮佳盈1号")
        self.assertEqual(rows[1]["我方账户"], "广粤尊享77号")
        self.assertEqual(rows[0]["交易规模万"], 1000)
        self.assertEqual(rows[1]["交易规模万"], 1000)
        self.assertEqual(rows[0]["原始净价"], "99.814")
        self.assertEqual(rows[1]["原始净价"], "99.814")


if __name__ == "__main__":
    unittest.main()
