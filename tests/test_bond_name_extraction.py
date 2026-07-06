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

    def test_table_style_rows_are_split(self):
        text = (
            "20260703 5+5 24科创债 某产品 123456.SH 1000 行权 2.2 99 某我方 3600001234 7012345\n"
            "20260703 5+5 24科创债 某产品2 123456.SH 2000 行权 2.3 100 某我方 3600001235 7012346"
        )
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["对方账户"], "某产品")
        self.assertEqual(rows[0]["行权收益率"], "2.2")
        self.assertEqual(rows[0]["约定号"], "7012345")
        self.assertEqual(rows[1]["对方账户"], "某产品2")
        self.assertEqual(rows[1]["行权收益率"], "2.3")
        self.assertEqual(rows[1]["约定号"], "7012346")

    def test_dealer_code_mapping_fallback_still_works(self):
        self.assertEqual(extractor.org_near("只有代码 000032 没有机构全称", "000032"), "广发证券")


if __name__ == "__main__":
    unittest.main()
