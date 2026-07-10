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
    def test_detect_template_routes_all_eight_shapes(self):
        samples = {
            "T1": "245555.SH 26中电K2 3000 净价100 中信信托信昱13号 出给 银河证券 约定号403",
            "T2": """520286.SZ 26临平Y1 1000 净价100 中信信托信昱13号
出给
交易商代码：000680
交易员代码：00IW0022
交易主体代码：3600003320
交易主体名称：中信建投证券自营
约定号16191619""",
            "T3": "283056.SH 26武发02 2000+2000+1000 净价100 中信信托信昱13号 出给 中信建投 约定号106+107+108",
            "T4": """245523.SH 山能YK02 2000 净价100 中信信托信昱13号 出给 财通证券
鹏富通达1号 1000 约定号123
鹏富通达2号 1000 约定号124""",
            "T5": """520314.SZ 26涪陵04 3000 净价100 粤财信托锐益1号 交易商号: 000032 交易员号：000W0007 交易主体代码：3600001825 交易主体名称：广发证券股份有限公司机构经纪
出给
交易商代码 000247华安证券
交易主体代码 3600063530华安资管湘赢36M039号集合资产管理计划
交易主体代码 3600064400华安资管众赢12M002号集合资产管理计划
交易主体代码 3600064439华安资管众赢12M003号集合资产管理计划
交易主体代码 3600064444华安资管众赢12M004号集合资产管理计划
交易主体代码 3600064442华安资管众赢12M005号集合资产管理计划
交易主体代码 3600064443华安资管众赢12M006号集合资产管理计划
交易员代码 006V0014蒋佳玮
约定号 19491949+19491950+19491951+19491952+19491953+19491954 2026-07-06 交易卖方发单""",
            "T6": """520309.SZ 26江滨03 10000 净价100 粤财信托锐益1号 交易商号: 000032 交易员号：000W0007 交易主体代码：3600001825 交易主体名称：广发证券股份有限公司机构经纪
出给
20260703 5+5 26江滨03 广东7号 520309.SZ 2000 行权 2.2 100 粤财信托锐益1号 3600002960 70300001
20260703 5+5 26江滨03 华能集团 520309.SZ 1000 行权 2.2 100 粤财信托锐益1号 3600002420 70300002
交易商代码：000262
交易主体代码：每行倒数第二个（3600开头）
交易员：007A0013 王玥
约定号：每行最后（703开头） 2026-07-03 交易买方发单""",
            "T7": """国海资管
281836.SH 26嘉亭01 3000 净价100 约定号123
281837.SH 26嘉亭02 2000 净价100 约定号124
中信信托华盈添利4号""",
            "T8": """520238.SZ 26黄控01 8000 净价100 中信信托信昱11号 出给 中信建投 约定号 19300930
【中信建投深交所要素】
交易商代码：000680
交易主体：3600003320 中信建投证券自营
交易员：00IW0022，刘思彤""",
        }
        for expected, text in samples.items():
            with self.subTest(expected=expected):
                self.assertEqual(extractor.detect_template(text), expected)

    def test_t4_grouped_yd_with_account_rows_is_not_downgraded_to_t3(self):
        text = """要素已定
 282471.SH 26天矿01 6000 净价100 中信信托信昱13号证券投资信托计划 Z06308
出给
平安信托泽鑫1 1000 Z06308
平安信托泽鑫8 1000 Z06308
平安信托泽鑫10 1000 Z06308
平安信托泽鑫20 1000 Z06308
中粮信托苏盈1 2000  中粮信托Z07506
 约定号 542+543+544+545 0422交易"""
        self.assertEqual(extractor.detect_template(text), "T4")
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 5)
        self.assertEqual([r["对方账户"] for r in rows], [
            "平安信托泽鑫1",
            "平安信托泽鑫8",
            "平安信托泽鑫10",
            "平安信托泽鑫20",
            "中粮信托苏盈1",
        ])
        self.assertEqual([r["交易规模万"] for r in rows], [1000, 1000, 1000, 1000, 2000])
        self.assertEqual([r["约定号"] for r in rows], ["542", "543", "544", "545", ""])
        self.assertEqual([r["过券"] for r in rows], ["平安信托", "平安信托", "平安信托", "平安信托", "中粮信托"])
        self.assertEqual(rows[-1]["备注"], "约定号原文未给出，需人工核对补充")
        self.assertTrue(all(r["交易日期"] == "2026-04-22" for r in rows))

    def test_mixed_alpha_chinese_bond_name_is_detected(self):
        text = """要素已定
520108.SZ 26GC云南绿能V1 10000 净价100 中信信托信昱13号证券投资信托计划 交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪
出给
深交:买方：中信证券股份有限公司机构经纪   交易商代码000262  交易主体代码：3600000001    交易员号：007A0001
 约定号 18480848 2026-05-11 交易卖方发单"""
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["债券简称"], "26GC云南绿能V1")
        self.assertEqual(rows[0]["对方账户"], "中信证券股份有限公司机构经纪")

    def test_compact_mmdd_plus_zero_header_keeps_bond_name(self):
        text = """要素已定
  0511+0  26GC云南绿能V1  520108.SZ  净价100   0.3e  广发证券（交易商：000262 交易员：007A0001 交易主体代码：3600000001 账户：中信信托信昱11号证券投资信托计划 ） to  富滇银行理财（账户：“富利添盈”周周赢开放式理财计划ZY2201期 本方交易商简称：广发证券 交易主体代码：3600001825 交易商代码：000032 交易员代码：000W0007）
约定号 10041004"""
        rows = extractor.parse_text(text)
        row = rows[0]
        self.assertEqual(row["债券简称"], "26GC云南绿能V1")
        self.assertEqual(row["交易日期"], "2026-05-11")
        self.assertEqual(row["清算速度"], "T+0")
        self.assertEqual(row["交易规模万"], 3000)
        self.assertEqual(row["对方账户"], "富利添盈周周赢开放式理财计划ZY2201期")

    def test_t2_buy_block_falls_back_to_broker_account_when_no_product_name(self):
        text = """要素已定
20260512+0  26淅川V1  520132.SZ  2000W  2.6281  净价100  中信信托（中信信托信昱13号证券投资信托计划）  to  华宝证券  约定号16391639

卖方（发）
交易商代码：000262
交易员代码：007A0001
交易主体代码/简称/量：
3600000001  中信证券股份有限公司机构经纪  2000


买方（点）
交易商代码：000640  华宝证券 
交易员号: 00HS0008 李哲人
交易主体代码：3600003979"""
        rows = extractor.parse_text(text)
        row = rows[0]
        self.assertEqual(row["交易方向"], "卖出")
        self.assertEqual(row["对方账户"], "华宝证券")
        self.assertEqual(row["过券"], "华宝证券")
        self.assertEqual(row["对手方交易员"], "李哲人")
        self.assertEqual(row["对手方交易商代码"], "000640")
        self.assertEqual(row["对手方交易主体代码"], "3600003979")
        self.assertEqual(row["交易日期"], "2026-05-12")
        self.assertEqual(row["清算速度"], "T+0")

    def test_t2_counterparty_head_strips_exchange_and_side_prefixes(self):
        text = """520108.SZ 26GC云南绿能V1 10000 净价100 中信信托信昱13号证券投资信托计划 交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪
出给
深交:  买方：中信证券华南，机构经纪   交易商代码000038，交易主体代码：3600003728  交易员号：00120002
 约定号 18480848 2026-05-12 交易买方发单"""
        rows = extractor.parse_text(text)
        row = rows[0]
        self.assertEqual(row["对方账户"], "中信证券华南,机构经纪")
        self.assertEqual(row["过券"], "中信证券华南")
        self.assertEqual(row["报价发起方"], "对方发起")

    def test_t3_counterparty_account_strips_zcode_and_split_sizes(self):
        text = """282631.SH 26吉投01 4000 净价100 中信信托信昱13号证券投资信托计划 Z06308
出给 首创证券Z08509 200+1000+1000
 约定号 100+101+102 2026-05-11 交易"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 3)
        self.assertEqual([r["对方账户"] for r in rows], ["首创证券", "首创证券", "首创证券"])
        self.assertEqual([r["过券"] for r in rows], ["首创证券", "首创证券", "首创证券"])
        self.assertEqual([r["交易规模万"] for r in rows], [200, 1000, 1000])
        self.assertEqual([r["约定号"] for r in rows], ["100", "101", "102"])

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

    def test_name_only_trade_can_parse_and_normalize_size(self):
        text = "要素已定\n26CHNG3K 净价100 0.1e 信昱13 出给 华润信托瑞合5号 中信Z06308 约定号 907"
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["债券代码"], "")
        self.assertEqual(rows[0]["债券简称"], "26CHNG3K")
        self.assertEqual(rows[0]["交易规模万"], 1000)
        self.assertEqual(rows[0]["我方账户"], "中信信托信昱13号")

    def test_counterparty_detail_block_does_not_become_name_only_trade(self):
        text = """买方（点）
交易商代码：000640  华宝证券
交易员号: 00HS0008 李哲人
交易主体代码：3600003979"""
        rows = extractor.parse_text(text)
        self.assertEqual(rows, [])

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

    def test_dealer_code_before_yd_is_not_misread_as_size(self):
        text = """要素已定
134978.SZ 26开城02 1000 净价100 中信信托信昱13号证券投资信托计划 交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪
出给
交易主体代码：3600056447
交易主体全称：方正证券鑫盛12号集合资产管理计划
交易员代码：毕晓韵（000S0008）
交易商代码 000028
约定号 10251025 2026-04-13 交易卖方发单"""
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["交易规模万"], 1000)
        self.assertEqual(rows[0]["对方账户"], "方正证券鑫盛12号集合资产管理计划")
        self.assertEqual(rows[0]["过券"], "方正证券")

    def test_k_and_kw_units_are_normalized_to_wan(self):
        samples = {
            "26伊犁02 净价100 2k 信昱13 出给 华创证券 约定号123": 2000,
            "26伊犁02 净价100 2kw 信昱13 出给 华创证券 约定号123": 2000,
            "26伊犁02 净价100 2000w 信昱13 出给 华创证券 约定号123": 2000,
        }
        for text, expected in samples.items():
            with self.subTest(text=text):
                rows = extractor.parse_text(text)
                self.assertEqual(rows[0]["交易规模万"], expected)

    def test_size_supports_thousand_separator_and_yuan_suffix(self):
        samples = {
            "245555.SH 26中电K2 1,000w 净价100 中信信托信昱13号 出给 银河证券 约定号403": 1000,
            "245555.SH 26中电K2 1000万元 净价100 中信信托信昱13号 出给 银河证券 约定号403": 1000,
            "245555.SH 26中电K2 0.1亿元 净价100 中信信托信昱13号 出给 银河证券 约定号403": 1000,
        }
        for text, expected in samples.items():
            with self.subTest(text=text):
                rows = extractor.parse_text(text)
                self.assertEqual(rows[0]["交易规模万"], expected)

    def test_comma_separated_single_line_still_recovers_unitless_size(self):
        text = "要素已定\n282248.SH,26菏发Z1,2500,净价100,上交所+0,约定号 959,中信信托信昱13号证券投资信托计划Z06308 出给 浦信恒兴2号私募基金Z07715"
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["债券简称"], "26菏发Z1")
        self.assertEqual(rows[0]["交易规模万"], 2500)
        self.assertEqual(rows[0]["清算速度"], "T+0")

    def test_fullwidth_and_lowercase_bond_code_are_normalized(self):
        text = "２４５５５５．ＳＨ　２６中电Ｋ２　０．１ｅ　净价１００　中信信托信昱13号　出给　银河证券　约定号４０３"
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["债券代码"], "245555.SH")
        self.assertEqual(rows[0]["交易规模万"], 1000)

        text2 = "245555.sh 26中电K2 0.1e 净价100 中信信托信昱13号 出给 银河证券 约定号403"
        rows2 = extractor.parse_text(text2)
        self.assertEqual(rows2[0]["债券代码"], "245555.SH")

    def test_bare_md_date_with_settlement_speed_is_parsed(self):
        # 真实样本里出现过的写法："6/30+0"/"4/14+0"：日期在前无交易所后缀，"+0"是清算速度
        text = ("6/30+0 281836.SH 26嘉亭01 3000 净价100 中信信托华盈添利4号 "
                "出给 国海资管 约定号123")
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["交易日期"], "2026-06-30")
        self.assertEqual(rows[0]["清算速度"], "T+0")

    def test_compact_yyyymmdd_date_is_parsed(self):
        # 表格式记录里常见的紧凑日期："20260703"，此前 RE_DATEC 定义了但没接入 norm_date
        text = "20260703 26江滨03 520309.SZ 2000 净价100 粤财信托锐益1号 出给 广东7号 约定号 70300001"
        rows = extractor.parse_text(text)
        self.assertEqual(rows[0]["交易日期"], "2026-07-03")

    def test_today_and_exchange_settlement_keywords(self):
        text1 = "今天+0 282907.SH 26浔产K2 1000 净价100 粤财信托添添益1号 to 财通证券 约定号208"
        rows1 = extractor.parse_text(text1)
        self.assertEqual(rows1[0]["清算速度"], "T+0")

        text2 = "282248.SH 26菏发Z1 2500 净价100 上交所+0 约定号959 中信信托信昱13号 出给 浦信恒兴2号私募基金"
        rows2 = extractor.parse_text(text2)
        self.assertEqual(rows2[0]["清算速度"], "T+0")

    def test_term_structure_plus_is_not_misread_as_settlement_speed(self):
        # "5+5" 是期限写法(5年+5年)，不是清算速度；不应被误读成 T+5
        text = "20260703\t5+5\t26江滨03\t广东7号\t520309.SZ\t2000\t\t行权\t2.2\t100\t粤财信托锐益1号\t3600002960\t70300001"
        rows = extractor.parse_text(text)
        self.assertTrue(all(r["清算速度"] != "T+5" for r in rows))

    def test_repeated_bond_code_with_name_before_code_keeps_name_code_aligned(self):
        # 简称写在代码前面（"26中财G3 245168.SH"）、且同一代码重复多行时，
        # 按代码字符位置切分会把简称错位到下一行；应改成按行首切分。
        text = """信昱11卖出
5/18  +1  3Y 26中财G3 245168.SH  1.73%/净价100  0.1e 中信信托【Z06308】 to 银华基金 【Z05400】 约定号929
5/18  +1  3Y 26中财G3 245168.SH  1.73%/净价100  0.1e 中信信托【Z06308】 to 银华基金 【Z05400】 约定号930
5/18  +1  3Y 26中财G3 245168.SH  1.73%/净价100  0.1e 中信信托【Z06308】 to 银华基金 【Z05400】 约定号931
5/18  +1  5Y 26中财G4 245169.SH  1.88%/净价100  0.3e 中信信托【Z06308】 to 银华基金 【Z05400】 约定号940"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 4)
        self.assertEqual(
            [(r["债券代码"], r["债券简称"]) for r in rows],
            [
                ("245168.SH", "26中财G3"),
                ("245168.SH", "26中财G3"),
                ("245168.SH", "26中财G3"),
                ("245169.SH", "26中财G4"),
            ],
        )
        self.assertEqual([r["交易规模万"] for r in rows], [1000, 1000, 1000, 3000])
        self.assertTrue(all(r["我方账户"] == "中信信托信昱11号" for r in rows))
        self.assertTrue(all(r["清算速度"] == "T+1" for r in rows))

    def test_no_direction_verb_record_still_recovers_counterparty_fields(self):
        # 整条记录完全没有"买入/卖出/出给/to/from"这类方向词，只在对方机构名后面标了个"发"字；
        # 之前 otherseg 会退化成空字符串，导致对方账户/过券/交易员/主体代码全部丢空
        text = """要素已定
2026/5/13	+0
520058.SZ	26焦资K1	净价100
首创证券	发
交易商代码：000613
交易主体代码：

3600061177	慧享17号24个月	1500  约定号  13161316


交易员号：00H10004  张跃曦

 交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual(r["交易规模万"], 1500)
        self.assertEqual(r["约定号"], "13161316")
        self.assertEqual(r["对方账户"], "慧享17号24个月")
        self.assertEqual(r["过券"], "首创证券")
        self.assertEqual(r["对手方交易员"], "张跃曦")
        self.assertEqual(r["对手方交易主体代码"], "3600061177")
        self.assertEqual(r["清算速度"], "T+0")
        self.assertEqual(r["报价发起方"], "对方发起")

    def test_multi_subject_block_without_any_yd_still_splits_rows(self):
        # 一段里有 2 个以上交易主体代码，但整段完全没有约定号——不能因此把两笔坍缩成一条空记录
        text = """2026/5/13	+0
520058.SZ	26焦资K1	净价100
首创证券	发
交易商代码：000613
交易主体代码：
3600060632	创赢PX02号	1700
3600059194 创赢PX01号 800

交易员号：00H10004  张跃曦

  交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 2)
        self.assertEqual([r["对方账户"] for r in rows], ["创赢PX02号", "创赢PX01号"])
        self.assertEqual([r["交易规模万"] for r in rows], [1700, 800])
        self.assertTrue(all(r["约定号"] == "" for r in rows))
        self.assertTrue(all("约定号" in r["备注"] for r in rows))

    def test_self_broker_identity_code_not_misread_as_counterparty_subject(self):
        # "交易主体名称：...机构经纪"后缀本身不是我方专属的——对手方自己的机构经纪代码
        # 也可能带这个后缀，但只会跟那一笔对手绑一次；真正能区分"这是我方固定占位码"的
        # 信号是重复：同一个交易商+交易员+主体代码三元组在同一份文本里对不同的我方产品/
        # 不同的记录反复出现，才该被当成我方占位码而不是对方自己的主体代码。
        text = """520230.SZ 26苏园GY02 2000 净价100 中信信托睿兴5号 交易商号: 000032 交易员号：000W0007 交易主体代码：3600001825 交易主体名称：广发证券股份有限公司机构经纪
出给
交易商：000262
交易商名称：中信证券股份有限公司
交易员：007A0001
交易员名称：中信证券机构经纪交易员
交易主体代码：3600000001
交易主体类型：03 机构经纪
交易主体名称：中信证券股份有限公司机构经纪
交易主体简称：中信证券机构经纪
 约定号 19440944 2026-06-18 交易卖方发单

134945.SZ 26武经Y1 3000 净价100 中信信托信昱13号证券投资信托计划
 交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪

出给
1.交易商代码：000001
2.交易商名称：国信证券股份有限公司
3.交易员代码：00010016
4.交易员名称：经纪1
5.交易主体代码：3600006811
6.交易主体名称：国信证券股份有限公司机构经纪
7.交易主体类型：机构经纪

 约定号 19590959 2026-03-27 交易卖方发单"""
        rows = extractor.parse_text(text)
        # 3600000001 在文本里重复出现(两条记录都报了同一个交易商+交易员+主体代码)，判定为我方占位码，排除
        self.assertEqual(rows[0]["对手方交易主体代码"], "")
        # 3600006811 只出现一次、绑定国信证券这一个对手，应该被当成对方真实的主体代码保留
        self.assertEqual(rows[1]["对手方交易主体代码"], "3600006811")

    def test_self_broker_block_in_middle_still_recovers_trailing_counterparty(self):
        # 我方经纪身份块("我方发"+交易商/交易员/主体代码四件套)夹在中间，真正的对方信息写在它后面，
        # 不是像之前的样本那样写在它前面——之前的兜底假设"块之后=我方段"会把对方信息一起划错
        text = """要素已定  我方发
 520128.SZ 26华旅Y1 1000 净价100 中信信托信昱13号证券投资信托计划
交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪
约定号 16261626

交易商代码 000681 国投证券
交易主体代码 3600062424
交易主体名称 国证资管建盈增利1号集合资产管理计划
交易员 00IX0005 郝爽"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual(r["我方账户"], "中信信托信昱13号")
        self.assertEqual(r["对方账户"], "国证资管建盈增利1号集合资产管理计划")
        self.assertEqual(r["过券"], "国投证券")
        self.assertEqual(r["对手方交易员"], "郝爽")
        self.assertEqual(r["对手方交易商代码"], "000681")
        self.assertEqual(r["对手方交易主体代码"], "3600062424")
        self.assertEqual(r["报价发起方"], "我方发起")

    def test_product_name_with_trailing_compound_suffix_is_not_truncated(self):
        # 产品名"N号"后面还接了一段描述性尾巴(集合资产管理计划)，非贪婪正则容易在"N号"处提前收尾
        text = "520128.SZ 26华旅Y1 1000 净价100 出给\n交易主体名称 国证资管建盈增利1号集合资产管理计划\n约定号123"
        self.assertEqual(extractor.pick_prod(text), "国证资管建盈增利1号集合资产管理计划")

    def test_single_line_broker_short_name_with_icode_still_fills_counterparty(self):
        text = "245535.SH 26平证12 3000 净价100 中信信托信昱13号 i020055109 出给 粤开证券i020059207 约定号 234 2026-06-30 交易"
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["对方账户"], "粤开证券")
        self.assertEqual(rows[0]["过券"], "粤开证券")

    def test_t8_supplement_enriches_metadata_without_overwriting_main_counterparty(self):
        text = """520238.SZ 26黄控01 8000 净价100 中信信托信昱11号 交易商号: 000262 交易员号：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪
出给  中信建投
约定号 19300930 2026-07-01 交易卖方发单
【中信建投深交所要素】
交易商代码：000680
交易主体：3600003320 中信建投证券自营
交易员：00IW0022，刘思彤"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["对方账户"], "中信建投")
        self.assertEqual(row["过券"], "中信建投证券")
        self.assertEqual(row["对手方交易商代码"], "000680")
        self.assertEqual(row["对手方交易员代码"], "00IW0022")
        self.assertEqual(row["对手方交易主体代码"], "3600003320")
        self.assertEqual(row["交易规模万"], 8000)
        self.assertEqual(row["交易日期"], "2026-07-01")

    def test_size_anchored_subject_blocks_pair_sizes_and_shared_yds(self):
        text = """要素已定
520160.SZ 26惠文Z1 3000 净价100 中信信托信昱11号证券投资信托计划 交易商：000262 交易员：007A0001 交易主体代码：3600000001 交易主体名称：中信证券股份有限公司机构经纪
出给
 1000w
交易主体：慧享8号
交易主体代码：3600025506
交易商代码：000025
交易员：耿涛

交易员代码：000P0032

1000w
交易主体：慧享20号24个月定开小集合
交易主体代码：3600054511
交易商代码：000025
交易员：耿涛
交易员代码：000P0032

1000w
交易主体：慧享11号18个月定开小集合
交易主体代码：3600049514
交易商代码：000025
交易员：耿涛
交易员代码：000P0032
 约定号 13441344+13551355+13661366 2026-05-14 交易卖方发单"""
        rows = extractor.parse_text(text)
        self.assertEqual(len(rows), 3)
        self.assertEqual([r["交易规模万"] for r in rows], [1000, 1000, 1000])
        self.assertEqual(
            [r["对方账户"] for r in rows],
            ["慧享8号", "慧享20号24个月定开小集合", "慧享11号18个月定开小集合"],
        )
        self.assertEqual([r["对手方交易主体代码"] for r in rows], ["3600025506", "3600054511", "3600049514"])
        self.assertEqual([r["约定号"] for r in rows], ["13441344", "13551355", "13661366"])
        self.assertTrue(all(r["过券"] == "光大证券" for r in rows))
        self.assertTrue(all(r["对手方交易商简称"] == "光大证券" for r in rows))

    def test_semantic_labels_support_unseen_accounts_and_explicit_values(self):
        text = """交易日期：2025年12月8日；清算速度：T+1；交易方向：买入
债券代码：SH 245555；债券简称：26中电K2；到期收益率：2.456；原始净价：99.880；交易净价：99.890；交易规模：1,500万
我方账户：华盈稳健新策略1号；对方账户：银河证券自营；约定号：998877"""
        row = extractor.parse_text(text)[0]
        expected = {
            "市场": "上交所",
            "交易方向": "买入",
            "交易日期": "2025-12-08",
            "债券代码": "245555.SH",
            "债券简称": "26中电K2",
            "到期收益率": "2.456",
            "原始净价": "99.880",
            "交易净价": "99.890",
            "交易规模万": 1500,
            "我方账户": "华盈稳健新策略1号",
            "对方账户": "银河证券自营",
            "清算速度": "T+1",
            "约定号": "998877",
        }
        for field, value in expected.items():
            with self.subTest(field=field):
                self.assertEqual(row[field], value)

    def test_bond_code_separator_and_prefix_variants_are_canonicalized(self):
        samples = {
            "245555-SH 26中电K2 1000万 净价100 中信信托信昱13号 出给 银河证券 约定号101": "245555.SH",
            "520286 SZ 26临平Y1 1000万 净价100 中信信托信昱13号 出给 中信建投证券 约定号102": "520286.SZ",
            "SH:245555 26中电K2 1000万 净价100 中信信托信昱13号 出给 银河证券 约定号103": "245555.SH",
        }
        for text, expected in samples.items():
            with self.subTest(text=text):
                row = extractor.parse_text(text)[0]
                self.assertEqual(row["债券代码"], expected)

    def test_from_direction_depends_on_our_account_position(self):
        buy = "中信信托信昱13号 245555.SH 26中电K2 1000万 净价100 FROM 银河证券自营 约定号101"
        sell = "银河证券自营 买入 245555.SH 26中电K2 1000万 净价100 FROM 中信信托信昱13号 约定号102"
        self.assertEqual(extractor.parse_text(buy)[0]["交易方向"], "买入")
        self.assertEqual(extractor.parse_text(sell)[0]["交易方向"], "卖出")

    def test_lowercase_trader_code_and_explicit_t_settlement_are_normalized(self):
        text = """2025年12月8日 520286-SZ 26临平Y1 1000万 净价100 中信信托信昱13号 出给 中信建投证券
交易商代码：000680
交易员代码：00iw0022 刘思彤
交易主体代码：3600003320
约定号：16191619
结算方式：T+0"""
        row = extractor.parse_text(text)[0]
        self.assertEqual(row["交易日期"], "2025-12-08")
        self.assertEqual(row["清算速度"], "T+0")
        self.assertEqual(row["对手方交易员代码"], "00IW0022")
        self.assertEqual(row["对手方交易员"], "刘思彤")

    def test_size_after_exercise_keyword_is_not_misread_as_exercise_yield(self):
        text = "2.05Y+NY 259348.SH 25北建Y2 AA+ 2.08行权 4000 今天交易所 中信信托信昱13号 出给 山西证券 约定号105"
        row = extractor.parse_text(text)[0]
        self.assertEqual(row["行权收益率"], "2.08")
        self.assertEqual(row["交易规模万"], 4000)

    def test_action_words_are_not_used_as_counterparty_accounts(self):
        text = """520285.SZ 26滁工K1 2000 净价100 中信信托华盈添利4号 交易商号:000680 交易员号:00IW0015 交易主体代码:3600003118 交易主体名称:中信建投证券股份有限公司机构经纪
出给
交易商:000680(中信建投证券)
交易主体:3600003118
交易员:00IW0013
约定号11061106 2026-07-02 交易卖方发单"""
        row = extractor.parse_text(text)[0]
        self.assertEqual(row["对方账户"], "中信建投证券")
        self.assertEqual(row["过券"], "中信建投证券")

    def test_party_blocks_override_inline_to_when_locating_counterparty(self):
        text = """①买入 4.96Y 520260.SZ 26文体02 100净价 2000W 07.02交易所 广发证券 to 世纪证券资管
买入方
交易商代码:000039 世纪证券
交易员号:00130016 李南阳
交易主体代码/简称/量/约定号:
3600062601 鑫享世成24M013号 500w 约定号12211221
3600063165 鑫享世成24M014号 500w 约定号13311331
卖出方
中信信托信昱11号
交易商号:000262
交易员号:007A0001
交易主体代码:3600000001
交易主体名称:中信证券股份有限公司机构经纪
卖家先发"""
        rows = extractor.parse_text(text)
        self.assertEqual([r["对方账户"] for r in rows], ["鑫享世成24M013号", "鑫享世成24M014号"])
        self.assertTrue(all(r["交易方向"] == "卖出" for r in rows))
        self.assertTrue(all(r["过券"] == "世纪证券" for r in rows))
        self.assertTrue(all(r["对手方交易商代码"] == "000039" for r in rows))

    def test_subject_account_boundary_excludes_yd_label(self):
        text = """134936.SZ 26肥产K1 2000 净价100 中信信托信昱13号证券投资信托计划
交易商:000262 交易员:007A0001 交易主体代码:3600000001 交易主体名称:中信证券股份有限公司机构经纪
出给
买入:国泰海通账户:
交易商代码:000612
交易员:赵越 00H00010
交易主体代码:3600060282
交易主体简称:国泰海通福星多元稳健1号集合资产管理计划(400万)约定号10061006
买入:国泰海通账户:
交易商代码:000612
交易员:赵越 00H00010
交易主体代码:3600060220
交易主体简称:国泰海通福星多元稳健2号集合资产管理计划(600万)约定号10061007"""
        rows = extractor.parse_text(text)
        self.assertEqual(
            [r["对方账户"] for r in rows],
            ["国泰海通福星多元稳健1号集合资产管理计划", "国泰海通福星多元稳健2号集合资产管理计划"],
        )
        self.assertTrue(all(r["过券"] == "国泰海通" for r in rows))

    def test_dealer_mapping_beats_product_fragment_for_guoquan(self):
        text = """520271.SZ 26韶旅01 2000 净价100 中信信托信昱13号 交易商号:000262 交易员号:007A0001 交易主体代码:3600000001 交易主体名称:中信证券股份有限公司机构经纪
出给
交易商代码:007128
交易主体代码:3600064069
交易主体名称:东方红添利38号集合资产管理计划
交易主体简称:东方红添利38号集合
交易员代码:05I00005
约定号18510851"""
        row = extractor.parse_text(text)[0]
        self.assertEqual(row["对方账户"], "东方红添利38号集合资产管理计划")
        self.assertEqual(row["过券"], "东方证券资管")

    def test_institution_head_is_account_while_dealer_is_guoquan(self):
        text = """520292.SZ 26交资K2 1000 净价100 中信信托信昱13号 交易商号:000262 交易员号:007A0001 交易主体代码:3600000001 交易主体名称:中信证券股份有限公司机构经纪
出给
贵阳银行:
交易商代码:000032 广发证券
交易主体:3600001825 广发证券机构经纪
交易员:000W0007
约定号19330933"""
        row = extractor.parse_text(text)[0]
        self.assertEqual(row["对方账户"], "贵阳银行")
        self.assertEqual(row["过券"], "广发证券")

    def test_unlabeled_known_dealer_code_can_supply_guoquan(self):
        text = """520312.SZ 26兴康02 2000 净价100 中信信托华盈添利4号 出给
合享3号 资管 000058 3600063778 华鑫证券合享3号集合资产管理计划
交易员代码:001M0033 约定号191"""
        row = extractor.parse_text(text)[0]
        self.assertEqual(row["对方账户"], "合享3号")
        self.assertEqual(row["过券"], "华鑫证券")

    def test_full_account_boundary_keeps_semantic_parentheses_and_internal_hyphen(self):
        cases = {
            "3600036691 山东省（拾号）职业年金计划－民生银行 1000万 约定号13520002":
                "山东省(拾号)职业年金计划",
            "3600063240 山西潞安矿业（集团）有限责任公司企业年金-中国工商银行 1000万 约定号13520003":
                "山西潞安矿业(集团)有限责任公司企业年金",
            "3600060000 方正证券建盈-玺福17号（家族专享）集合资产管理计划（400万） 约定号13520004":
                "方正证券建盈-玺福17号(家族专享)集合资产管理计划",
        }
        for line, expected in cases.items():
            with self.subTest(line=line):
                self.assertEqual(extractor.extract_full_account_from_line(line), expected)

    def test_custodian_suffix_is_removed_only_after_complete_pension_account(self):
        line = "3600057500 天津市拾号职业年金计划－中信银行股份有限公司 1000万 约定号13520001"
        self.assertEqual(
            extractor.extract_full_account_from_line(line),
            "天津市拾号职业年金计划",
        )


if __name__ == "__main__":
    unittest.main()
