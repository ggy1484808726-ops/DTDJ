# -*- coding: utf-8 -*-
"""
现券预录单要素识别 · 正则提取脚本
思路：切段(我方/对方) + 白名单认领我方 + 字段正则 + 拆单展开
用法：python3 提取脚本.py  → 生成 现券预录单要素_脚本输出.xlsx
"""
import re, glob, os

OUTPUT_COLUMNS = [
    "市场","交易方向","交易日期","债券代码","债券简称","到期收益率","行权收益率",
    "原始净价","交易净价","交易规模万","我方账户","对方账户","过券","中介","中介费",
    "对手方交易员","清算速度","约定号","对手方交易员代码","对手方交易商代码",
    "对手方交易商简称","对手方交易主体代码","报价发起方","备注","原文"
]

# ========== 1. 我方账户白名单（按实际管理产品维护）==========
WHITELIST = [
    "中信信托信昱11号","中信信托信昱13号","中信信托华盈添利4号",
    "粤财信托添添益1号","粤财信托锐益2号","粤财信托广盈恒益1号",
    "财信信托湘信财盈73号","财信信托湘信财盈74号",
]

# ========== 2. 字段正则 ==========
RE_BOND   = re.compile(r'(\d{6})\.(SH|SZ)')
RE_PRICE  = re.compile(r'净价\s*(\d{2,3}(?:\.\d+)?)')
RE_PRICE2 = re.compile(r'(?<!\d)(1\d{2}(?:\.\d+)?)(?!\d)')      # 兜底 100~199
RE_YIELD  = re.compile(r'(\d+(?:\.\d+)?)\s*%')
RE_YIELD2 = re.compile(r'(?<![\d.])([12]\.\d{2,3})(?![\d%])')   # 裸 1.xx/2.xx
RE_DATEF  = re.compile(r'(20\d{2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})')
RE_DATES  = re.compile(r'(?<!\d)(\d{1,2})[.](\d{1,2})\s*交易所')
RE_YD     = re.compile(r'约(?:定号)?[：:\s]*([0-9]+(?:\s*[+＋、，,]\s*[0-9]+)*)')  # 支持 + 、 ，分隔及"约"简写
RE_YD_ALL = re.compile(r'\b(\d{3,10})\b')
RE_DEALER = re.compile(r'交易商(?:代码|号)?\s*[:：]?\s*(\d{6})')
RE_TRADER = re.compile(r'交易员(?:代码|号)?\s*[:：]?\s*([0-9A-Z]{8})')
RE_TRADER2= re.compile(r'交易员(?:及交易员代码|代码|名称|号)?\s*[：:]?\s*[一-龥]{2,4}\s*[（(]?\s*([0-9A-Z]{7,8})')  # 名在码前：赵越 00H00010 / 付玉 007Z0039
RE_SUBJ   = re.compile(r'(?:交易(?:商)?主体(?:代码)?)\s*[:：]?\s*(36\d{8})')
RE_SUBJNM = re.compile(r'交易主体名称\s*[:：]?\s*([一-龥A-Z0-9\-（）()]+)')
RE_ICODE  = re.compile(r'i\d{9}')
RE_PERSON = re.compile(r'[（(]([一-龥]{2,4})[)）]')
RE_ORG    = re.compile(r'([一-龥]{2,10}(?:证券资管|证券自营|证券机构经纪|证券|资管|基金|信托|银行|期货|养老|保险|人寿|投顾|财富|资产))')
RE_PROD   = re.compile(r'([一-龥A-Za-z0-9]{2,18}?(?:\d+M?\d*号|\d+号))')
RE_INIT   = re.compile(r'(卖方发单|买方发单|卖家先发|买家先发|卖出先发|买入先发)')
RE_MID    = re.compile(r'中介费\s*[:：]\s*(\d+(?:\.\d+)?)')
RE_SIZEU  = re.compile(r'(\d+(?:\.\d+)?)\s*([wWkK万])')

MKT = {"SH":"上交所","SZ":"深交所"}
# 我方产品模式（白名单兜底）：中信/粤财/财信信托…号
RE_MINE_PAT = re.compile(r'((?:中信信托|粤财信托|财信信托)[一-龥A-Za-z0-9]{1,12}号)')

RE_SIZE_NJ = re.compile(r'(?<![.\d])(\d{3,6})\s*净价')          # 3000 净价100（前不接小数点，避免2.025误取025）
RE_SIZE_NJ2= re.compile(r'(\d{3,6})\s+\d+(?:\.\d+)?\s*净价')  # 5000 1.98 净价
RE_SIZE_PR = re.compile(r'(\d{3,6})\s+1\d{2}(?:\.\d+)?(?!\d)')  # 1000 100.001
RE_DEALER_ORG = re.compile(r'交易商(?:代码|号)?\s*[:：]?\s*\d{6}\s*[（(]?\s*([一-龥]{2,10}(?:证券资管|证券自营|证券机构经纪|证券|资管|基金|信托|银行|期货))')

def get_size(rec):
    m=re.search(r'[\s，,、](\d{2,6})\s*约定号', rec)   # 规模紧接约定号前(如 26武发02 2000 约定号857)；i码在前会被空格约束排除
    if m: return int(m.group(1))
    m=re.search(r'(?<![.\d])(\d{3,6})\s+\d\.\d{2,4}', rec)  # 规模+收益率(如 5000 2.025 净价)
    if m: return int(m.group(1))
    m=re.search(r'(?:行权|到期)\s*(\d{3,6})', rec)          # 2.08行权 4000
    if m: return int(m.group(1))
    m=RE_SIZE_NJ.search(rec)
    if m: return int(m.group(1))
    m=RE_SIZE_NJ2.search(rec)
    if m: return int(m.group(1))
    m=RE_SIZEU.search(rec)
    if m:
        v=float(m.group(1)); u=m.group(2).lower()
        return int(v*1000) if u=='k' else int(v)
    m=RE_SIZE_PR.search(rec)
    if m: return int(m.group(1))
    return ""

BADNAME_WORDS = (
    "净价", "交易", "约定", "出给", "买入", "卖出", "交易所",
    "主体", "名称", "代码", "票面", "到期", "行权", "发单", "先发", "报价",
)
# 债券简称：仅抓代码前后紧邻的债券命名形态，避免把 i 码尾号 / 交易动词误当简称
NAMEPATS = [
    re.compile(r'(?<![A-Za-z0-9一-龥])(\d{2}\s*[一-龥]{1,6}\s*[A-Z]{1,3}\s*\d{1,3})(?![A-Za-z0-9一-龥])'),
    re.compile(r'(?<![A-Za-z0-9一-龥])(\d{2}\s*[一-龥]{1,6}\s*\d{2})(?![A-Za-z0-9一-龥])'),
    re.compile(r'(?<![A-Za-z0-9一-龥])([一-龥]{2,4}\s*[A-Z]{1,3}\s*\d{2,3})(?![A-Za-z0-9一-龥])'),
    re.compile(r'(?<![A-Za-z0-9一-龥])([一-龥]{1,3}\s*\d{2}\s*[一-龥]{1,3}\s*[A-Z]{1,3}\s*\d{1,3})(?![A-Za-z0-9一-龥])'),
]
def _clean_name(s):
    nm=re.sub(r'\s+','',s)
    if not nm or not re.search(r'[一-龥]', nm): return ""
    if any(word in nm for word in BADNAME_WORDS): return ""
    return nm if 4 <= len(nm) <= 12 else ""
def _pick_name(seg, prefer='last'):
    hits=[]
    for pat in NAMEPATS:
        for m in pat.finditer(seg):
            nm=_clean_name(m.group(1))
            if nm:
                hits.append((m.start(1), m.end(1), nm))
    if not hits: return ""
    if prefer=='first':
        hits.sort(key=lambda x: (x[0], -(x[1]-x[0])))
    else:
        hits.sort(key=lambda x: (-x[1], -(x[1]-x[0]), x[0]))
    return hits[0][2]
def get_name(rec, code):
    i=rec.find(code); L=len(code)
    before=rec[max(0,i-24):i]; after=rec[i+L:i+L+24]
    # 优先代码后的简称，避免把代码前的 i 码尾号 / 买卖动词拼成伪简称
    name=_pick_name(after, prefer='first')
    if name: return name
    return _pick_name(before, prefer='last')

def dealer_org(seg):
    m=RE_DEALER_ORG.search(seg)
    return m.group(1) if m else ""

# 多产品拆单：返回 [(对方账户,主体代码,规模万,约定号,交易员码)] ；<2 个则空
# 上交所多拆：逐行"产品名(可无号) [可选i码] 规模(单位可选) 约定号"；规模须空格分隔+紧跟约定号，避免误匹配i码/单笔
RE_MULTI_SH = re.compile(r'(?:^|\n)[ \t]*([一-龥A-Za-z0-9][^\n]*?)(?:i\d{9})?[ \t（(]+(\d{1,6})\s*([wWkK万]?)[）)]?\s*[；;，, ]*约定号[：:\s]*(\d+)', re.M)
# 上交所多拆(规模在前)：如 "1）1500万 广东省拾贰号职业年金计划海富通组合 约定号842"
RE_MULTI_SH2= re.compile(r'(?:^|\n)\s*(?:\d{1,2}\s*[）)、.]\s*)?(\d{2,6})\s*([wWkK万])\s*([一-龥A-Za-z0-9（）()－\-]+?)\s*约定号[：:\s]*(\d+)', re.M)
RE_MULTI_SZ = re.compile(r'(36\d{8})\s*([^\d\s：:][^\n]*?)\s*(?:[（(]?\s*(\d+)\s*[wW万][）)]?)?\s*约定号[：:\s]*(\d+)')
RE_SUBJ_LABEL = re.compile(r'^(?:交易)?(?:商)?主体(?:名称|简称|代码)?\s*[：:]\s*')  # 去名字前的标签
RE_MULTI_SZ2= re.compile(r'([一-龥A-Za-z0-9]{2,16}号)\s*(\d+)\s*[wW万]\s*约定号\s*(\d+)')
RE_MULTI_FZ = re.compile(r'([一-龥]{2,4}\d{1,2})\s*(\d+)\s*万[；;，, ]*约定号\s*(\d+)')   # 玺福19 500万 约定号18420842
RE_SUBJ_NAME= re.compile(r'(36\d{8})\s*[^0-9]{0,4}?([一-龥A-Za-z0-9\-（）]+号)')            # 主体代码+全名
def extract_multi(seg, mkt):
    out=[]
    if mkt=='上交所':
        for m in RE_MULTI_SH.finditer(seg):        # 产品名在前
            v=int(m.group(2)); u=m.group(3).lower()
            sz=v*1000 if u=='k' else v
            nm=RE_SUBJ_LABEL.sub('', m.group(1)).strip(' \t：:，,；;')
            out.append((nm,'',sz,m.group(4),''))
        if len(out)<2:                              # 规模在前(如 1）1500万 广东省…组合 约定号842)
            for m in RE_MULTI_SH2.finditer(seg):
                v=int(m.group(1)); u=m.group(2).lower()
                sz=v*1000 if u=='k' else v
                nm=RE_SUBJ_LABEL.sub('', m.group(3)).strip(' \t：:，,；;（(').strip()
                out.append((nm,'',sz,m.group(4),''))
    else:
        for m in RE_MULTI_SZ.finditer(seg):
            sz=int(m.group(3)) if m.group(3) else ''
            nm=RE_SUBJ_LABEL.sub('', m.group(2)).strip(' \t：:，,；;（(').strip()
            out.append((nm,m.group(1),sz,m.group(4),''))
        if len(out)<2:   # 方正式：短名+规模万+约定号 分列，主体代码/全称另行分列(标签不一)
            fz=RE_MULTI_FZ.findall(seg)
            if len(fz)>=2:
                # 全称按位置配"其前最近的36主体代码"
                codes=[(m.start(),m.group(1)) for m in re.finditer(r'(36\d{8})', seg)]
                names=[(m.start(),m.group(1)) for m in re.finditer(r'(?:全称|名称)[：:]?\s*([一-龥][^；;\n，]*?号)', seg)]
                full2code={}
                for npos,nm in names:
                    prev=[c for cpos,c in codes if cpos<npos]
                    if prev: full2code[nm]=prev[-1]
                out=[]
                for short,size,yd in fz:
                    subj=''; acct=short+'号'
                    for full,c in full2code.items():
                        if short in full:
                            subj=c; acct=full; break
                    out.append((acct,subj,int(size),yd,''))
        if len(out)<2:   # 主体代码+名+规模逐行，约定号在末尾"+"连写(如东吴证券多产品)
            prods=re.findall(r'(36\d{8})\s*([^\d\n][^\n]*?)\s*(\d+)\s*[wW万]', seg)
            ydm=re.search(r'约定号[：:\s]*([0-9]+(?:[+＋][0-9]+)+)', seg)
            if len(prods)>=2 and ydm:
                yds=re.split(r'[+＋]', ydm.group(1))
                if len(yds)==len(prods):
                    out=[(RE_SUBJ_LABEL.sub('',n).strip(' \t：:，,；;（('), c, int(s), yds[i], '')
                         for i,(c,n,s) in enumerate(prods)]
    # 规模在独立行(如临动GT01: 2000\n交易商代码…\n约定号)——产品都缺规模时，按顺序配独立数字行
    if out and all(not x[2] for x in out):
        szl=re.findall(r'(?:^|\n)[ \t]*(\d{3,6})[ \t]*万?[ \t]*(?=\n)', seg)
        if len(szl)>=len(out):
            out=[(a,b,int(szl[i]),d,e) for i,(a,b,c,d,e) in enumerate(out)]
    return out if len(out)>=2 else []

def extract_paren_products(rec):
    """括号内'+'连写的产品串：(西藏信托 长盈稳健25号1200+西藏信托 善盈长盈稳健42号1600+...) → [(名,规模)]"""
    m=re.search(r'[（(]([^）)]*[一-龥][^）)]*[+＋][^）)]*)[）)]', rec)
    if not m: return []
    out=[]
    for it in re.split(r'[+＋]', m.group(1)):
        mm=re.match(r'\s*([一-龥].*?)\s*(\d{2,5})\s*$', it)
        if not mm: return []          # 有一项不符合(名+规模)就整体放弃
        out.append((mm.group(1).strip(), int(mm.group(2))))
    return out if len(out)>=2 else []

def extract_multibond(rec):
    """同一对手、多只券（每只券各带规模/净价/收益率/约定号）→ 逐券一行"""
    codes=[(m.start(),m.group(0)) for m in RE_BOND.finditer(rec)]
    if len(codes)<2: return []
    out=[]
    for idx,(pos,code) in enumerate(codes):
        end=codes[idx+1][0] if idx+1<len(codes) else len(rec)
        chunk=rec[pos:end]
        ydm=RE_YD.search(chunk)
        if not ydm: return []          # 不是"每券各自约定号"结构
        ym=RE_YIELD.search(chunk)
        pm=RE_PRICE.search(chunk) or RE_PRICE2.search(chunk.replace(code,''))
        out.append(dict(code=code,name=get_name(rec,code),size=get_size(chunk),
            yld=(ym.group(1)+'%') if ym else '',price=pm.group(1) if pm else '',yd=ydm.group(1)))
    return out if len(out)>=2 else []


def norm_date(rec):
    m = RE_DATEF.search(rec)
    if m: return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = RE_DATES.search(rec)
    if m: return f"2026-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    return ""

def find_mine(rec):
    """返回 (我方账户, 命中位置) ；没命中返回(None,-1)"""
    best=None; pos=-1
    for w in WHITELIST:
        i=rec.find(w)
        if i>=0 and (pos<0 or i<pos): best,pos=w,i
    if best: return best,pos
    # 兜底1：核心词
    for w in WHITELIST:
        core=re.sub(r'^(中信信托|粤财信托|财信信托)','',w)
        i=rec.find(core)
        if i>=0 and (pos<0 or i<pos): best,pos=w,i
    if best: return best,pos
    # 兜底2：模式识别我方产品（中信/粤财/财信信托…号），对方多为银行/基金/券商/年金，不会误匹配
    m=RE_MINE_PAT.search(rec)
    if m: return m.group(1), m.start()
    return best,pos

def direction_and_split(rec, mine_pos):
    """判断方向，返回(方向, 我方段, 对方段)"""
    # 显式买卖方标签：卖方/卖出方、买方/买入方
    ms=re.search(r'卖(?:出)?方', rec); mb=re.search(r'买(?:入)?方', rec)
    if ms and mb:
        lo,hi=sorted([ms.start(), mb.start()])
        seg1,seg2=rec[lo:hi],rec[hi:]
        if lo<=mine_pos<hi: mineseg,otherseg=seg1,seg2
        else: mineseg,otherseg=seg2,seg1
        direction="卖出" if "卖" in mineseg[:4] else "买入"
        return direction, mineseg, otherseg
    # 分隔符 出给/to
    for sep in ["出给"," to ","to "]:
        j=rec.find(sep)
        if j>=0:
            left,right=rec[:j],rec[j+len(sep):]
            return ("卖出",left,right) if mine_pos<j else ("买入",right,left)
    # from：X 买入 ... from 我方 → 我方=卖方
    j=rec.find("from")
    if j>=0:
        return "卖出", rec[j+4:], rec[:j]
    # 对方买入 …… 我方(在后)：如"金融街证券 i… 买入 bond … 中信信昱11号" → 我方卖出，对方在买入前
    j=rec.find("买入")
    if j>=0 and rec[j:j+3]!="买入方":
        return ("卖出", rec[j+2:], rec[:j]) if mine_pos>j else ("买入", rec[:j], rec[j+2:])
    return "", rec, ""

def org_near(rec, code6):
    """按6位交易商代码就近取机构名（前后各12字），如 '首创证券 交易商代码：000613' / '交易商代码：000039 世纪证券'"""
    if not code6: return ""
    i=rec.find(code6)
    if i<0: return ""
    for region in (rec[i+6:i+6+14], rec[max(0,i-14):i]):
        m=RE_ORG.search(region)
        if m: return m.group(1)
    return ""

def other_codes(rec, mine, mine_pos):
    """取"非我方"的交易商/交易员代码：我方代码紧跟白名单之后，其余即对手方"""
    tail = rec[mine_pos: mine_pos+120] if (mine_pos is not None and mine_pos>=0) else ""
    md = RE_DEALER.search(tail); md = md.group(1) if md else ""
    mt = RE_TRADER.search(tail); mt = mt.group(1) if mt else ""
    ds=[m.group(1) for m in RE_DEALER.finditer(rec) if m.group(1)!=md]
    ts=[m.group(1) for m in RE_TRADER.finditer(rec) if m.group(1)!=mt]
    if not ts:   # 兜底1：名在码前（交易员：赵越 00H00010）
        ts=[m.group(1) for m in RE_TRADER2.finditer(rec) if m.group(1)!=mt]
    if not ts:   # 兜底2：括号内的交易员代码（毕晓韵 （000S0008））
        ts=[c for c in re.findall(r'[（(]([0-9A-Z]{7,8})[)）]', rec) if c!=mt]
    return (ds[0] if ds else ""), (ts[0] if ts else "")

RE_NOTNAME=re.compile(r'券|经纪|交易|证券|机构|资管|基金|信托|银行|资产|公司|约定|集合|计划|组合|年金|产品|管理|保险|养老|席位|主体|代码|名称|净价|类型|资本|银|买|卖|发|出给|存续|持有')
def person_name(seg):
    """对手方交易员姓名：名可在码前/码后/单独'交易员：名'，排除'…机构经纪交易员'等非人名"""
    pats=[
        r'交易员\s*(?:代码|号)?\s*[：:]?\s*i\d{9}\s+([一-龥]{2,4})(?![一-龥])',            # 上交所i码后名：i029904205 孙磊
        r'交易员[号代码]*\s*[：:]\s*[0-9A-Z]{7,8}\s*[，,、\s]+([一-龥]{2,4})(?![一-龥])',  # 码 后名：04SE0003 周玲玲
        r'交易员(?:及交易员代码|代码|名称|号)?\s*[：:]?\s*([一-龥]{2,4})\s*[（(]?\s*[0-9A-Z]{7,8}',  # 名 后码：毕晓韵（000S0008）
        r'交易员\s*名称\s*[：:]\s*([一-龥]{2,4})(?![一-龥])',                                  # 交易员名称：张锐
        r'交易员\s*[：:]\s*([一-龥]{2,4})(?![一-龥0-9A-Z])',                                   # 交易员：谭骅
        r'i\d{9}\s+([一-龥]{2,4})(?![一-龥])',                                                # 无标签 i码后名：i020038603 吴静静
    ]
    for p in pats:
        for m in re.finditer(p, seg):
            nm=m.group(1)
            if not RE_NOTNAME.search(nm): return nm
    m=re.search(r'[（(]([一-龥]{2,4})[)）]', seg)
    if m and not RE_NOTNAME.search(m.group(1)): return m.group(1)
    return ""

def pick_org(seg):
    m=RE_ORG.search(seg)
    return m.group(1) if m else ""
def pick_prod(seg):
    m=RE_PROD.search(seg)
    return m.group(1) if m else ""

def expand_splits(rec):
    """返回 [(规模万, 约定号)] 列表；支持 + 连写"""
    ydm=RE_YD.search(rec)
    yds=[]
    if ydm:
        yds=[x.strip() for x in re.split(r'[+＋、，,]', ydm.group(1))]
    # 规模：找 a+b+c（含括号内 (1000+1000)）
    sizes=[]
    plus=re.search(r'(\d+(?:\s*[+＋]\s*\d+)+)', rec)
    if plus:
        sizes=[int(x) for x in re.split(r'[+＋]', plus.group(1))]
    return yds, sizes

def parse_record(rec):
    rec=re.sub(r'[ \t]+',' ',rec.strip())
    bm=RE_BOND.search(rec)
    if not bm: return []
    code=bm.group(0); mkt=MKT[bm.group(2)]
    # 简称
    tail=rec[bm.end():bm.end()+16]
    nm=re.search(r'\d{2}[一-龥]{1,5}[A-Z0-9]{0,4}', bm.group(0).join(['',''])+' '+tail)
    name=get_name(rec, code)
    date=norm_date(rec)
    price=""
    pm=RE_PRICE.search(rec)
    if pm: price=pm.group(1)
    else:
        pm2=RE_PRICE2.search(rec.replace(code,''))
        if pm2: price=pm2.group(1)
    ym=RE_YIELD.search(rec); yld = ym.group(1)+'%' if ym else ""
    xk=re.search(r'(\d+(?:\.\d+)?)\s*行权', rec); xingquan = xk.group(1) if xk else ""  # 2.08行权→行权收益率2.08
    dqk=re.search(r'(\d+(?:\.\d+)?)\s*到期', rec)
    if dqk and not yld: yld = dqk.group(1)
    init_raw=RE_INIT.search(rec); init_raw=init_raw.group(1) if init_raw else ""
    mid=RE_MID.search(rec); midfee = mid.group(1) if mid else ""

    mine,mpos=find_mine(rec)
    direction, mineseg, otherseg = direction_and_split(rec, mpos if mpos>=0 else 0)
    # 方向覆盖：看我方白名单附近的显式动词（应对我方在句尾且带"卖出/卖出先发"的情形，如"广发 to 首创…中信信昱11号 卖出"）
    if mine and mpos>=0:
        win=rec[mpos:mpos+90]
        if re.search(r'卖出|卖方|卖家', win): direction='卖出'
        elif re.search(r'买入方|买家', win): direction='买入'
    if not mine: mine="(未命中白名单)"

    # 报价发起方归一
    initiator=""
    if init_raw:
        seller_first = init_raw in ("卖方发单","卖家先发","卖出先发")
        if direction=="卖出": initiator="我方发起" if seller_first else "对方发起"
        elif direction=="买入": initiator="对方发起" if seller_first else "我方发起"

    # 对方字段（对方段）
    o_dealer = (RE_DEALER.search(otherseg) or [None,""])
    o_dealer = RE_DEALER.search(otherseg).group(1) if RE_DEALER.search(otherseg) else ""
    o_trader = (RE_TRADER.search(otherseg) or RE_TRADER2.search(otherseg))
    o_trader = o_trader.group(1) if o_trader else ""
    o_subj   = RE_SUBJ.search(otherseg).group(1) if RE_SUBJ.search(otherseg) else ""
    o_icode  = RE_ICODE.search(otherseg).group(0) if RE_ICODE.search(otherseg) else ""
    o_person = person_name(otherseg)
    o_org    = dealer_org(otherseg) or pick_org(otherseg)
    o_prod   = pick_prod(otherseg)
    # 对方主体名称/简称（全名，优先作对方账户）
    msn=re.search(r'(?:主体名称|主体简称)\s*[：:]\s*([一-龥A-Za-z0-9－\-（）()·]{3,60})', otherseg)
    o_subjname = msn.group(1).strip() if msn else ""
    if not o_subjname:  # 兜底：交易主体代码（宝惠）式，括号内为对方简称
        mp=re.search(r'交易主体代码[（(]([一-龥A-Za-z0-9]{2,12})[)）]', otherseg)
        if mp: o_subjname=mp.group(1).strip()
    counter_dealer_code = o_dealer                      # 对手方交易商代码(深)
    counter_trader_code = o_trader if mkt=="深交所" else o_icode  # 对手方交易员代码
    counter_subj = o_subj
    counter_short= o_org                                # 对手方交易商简称
    counter_acct = o_subjname or o_prod or o_org        # 对方账户：主体全名 > 产品 > 机构
    if mkt=="上交所":                                    # 上交所无交易商三件套 → 强制空
        counter_dealer_code=""; counter_short=""; counter_subj=""
    guoquan = counter_short                             # 过券=对手方交易商
    if not guoquan: guoquan=counter_acct

    # 净价
    orig=price; deal=price
    if midfee and price:
        try:
            p=float(price); f=float(midfee)
            deal=str(p-f) if direction=="卖出" else str(p+f)
        except: pass

    multibond=extract_multibond(rec)
    multi=extract_multi(rec, mkt)
    yds,sizes=expand_splits(rec)
    rows=[]
    def mk(size, yd, acct=None, subj=None):
        return dict(市场=mkt,交易方向=direction,交易日期=date,债券代码=code,债券简称=name,
            到期收益率=yld,行权收益率=xingquan,原始净价=orig,交易净价=deal,交易规模万=size,
            我方账户=mine,对方账户=acct if acct else counter_acct,过券=guoquan,中介="无" if not midfee else "",
            中介费=midfee if midfee else "无",对手方交易员=o_person,清算速度="",约定号=yd,
            对手方交易员代码=counter_trader_code,对手方交易商代码=counter_dealer_code,
            对手方交易商简称=counter_short,对手方交易主体代码=subj if subj else counter_subj,
            报价发起方=initiator,原文=rec,备注="")
    if multibond:                       # 同对手多只券：逐券一行
        for b in multibond:
            r=mk(b['size'], b['yd'])
            r['债券代码']=b['code']; r['债券简称']=b['name']
            r['到期收益率']=b['yld']; r['原始净价']=b['price']; r['交易净价']=b['price']
            rows.append(r)
        return rows
    paren=extract_paren_products(rec)   # 括号内'+'连写产品：过券取机构、对方账户取去机构后的产品名
    if paren:
        org=pick_org(otherseg) or pick_org(rec)
        for i,(nm,sz) in enumerate(paren):
            acct=re.sub(r'^'+re.escape(org)+r'\s*','',nm) if org else nm
            r=mk(sz, yds[i] if i<len(yds) else "", acct=acct)
            r["过券"]= org or acct
            rows.append(r)
        return rows
    if multi:
        odc,otc=other_codes(rec, mine, mpos)
        short = org_near(rec, odc) if mkt=="深交所" else ""
        sh_tcode=""                       # 上交所共享i码(对方段，排除我方)
        if mkt=="上交所":
            my_i=RE_ICODE.search(mineseg)
            my_i=my_i.group(0) if my_i else ""
            for m in RE_ICODE.finditer(otherseg):
                if m.group(0)!=my_i: sh_tcode=m.group(0); break
        # 若各产品规模都缺、但有总额且能整除份数 → 按份数均分
        even=""
        if all(not x[2] for x in multi):
            tot=get_size(rec)
            if isinstance(tot,int) and tot % len(multi)==0: even=tot//len(multi)
        for acct,subj,sz,yd,tcode in multi:
            r=mk(sz,yd,acct=acct,subj=subj)
            r["对手方交易员代码"]= tcode or sh_tcode or otc
            if mkt=="深交所":
                if odc: r["对手方交易商代码"]=odc
                if short: r["对手方交易商简称"]=short
            r["过券"]= short or pick_org(otherseg) or acct   # 过券优先机构
            if not sz and even:
                r["交易规模万"]=even; r["备注"]="规模原文未分列，按总额均分，请核"
            elif not sz:
                r["备注"]="规模未在原文分列，需人工确认"
            rows.append(r)
        return rows
    if sizes and yds and len(sizes)==len(yds):
        for s,y in zip(sizes,yds): rows.append(mk(s,y))
    elif len(yds)>1:
        # 多约定号但规模不是+连写：整单，逐约定号一行，规模留总额提示
        for y in yds: rows.append(mk("",y))
        for r in rows: r["备注"]="拆单-规模/对方账户需人工核对"
    else:
        size=get_size(rec)
        rows.append(mk(size, yds[0] if yds else ""))
    return rows

def split_records(text):
    raw=[]
    for chunk in re.split(r'\n\s*\n', text):
        if RE_BOND.search(chunk):
            raw.append(chunk)
        elif raw and chunk.strip():        # 无债券代码=上一条的续接(约定号/卖出方等)，用换行连接保留结构
            raw[-1]+='\n'+chunk
    # 二次切
    recs=[]
    for chunk in raw:
        lines=[l for l in chunk.split('\n') if l.strip()]
        # a) 每行都自成一笔(各含债券代码+约定号)→按行拆(如上交所两笔挨着无空行)
        if len(lines)>1 and all(RE_BOND.search(l) and RE_YD.search(l) for l in lines):
            recs.extend(lines); continue
        # b) 块内多个"各自带出给/买入的完整交易"(方向标记在每段内)→按债券代码切；
        #    "共享抬头+多只券"(段内无方向标记，如 华创 出给 财信 后跟两只券)不切，留给 multibond
        codes=[mm.start() for mm in RE_BOND.finditer(chunk)]
        if len(codes)>=2:
            segs=[chunk[codes[i]:codes[i+1]] for i in range(len(codes)-1)]+[chunk[codes[-1]:]]
            if all(re.search(r'出给|\bto\b|from|买入|卖出|卖方|买方', s) for s in segs):
                recs.extend(segs); continue
        recs.append(chunk)
    return recs

def parse_text(text):
    allrows=[]
    for rec in split_records(text or ""):
        allrows.extend(parse_record(rec))
    # 未命中白名单的行（多为对手只报券商席位、我方仅以席位i码出现）→ 打备注提示人工确认
    for r in allrows:
        if str(r.get("我方账户","")).startswith("(未") and not r.get("备注"):
            r["备注"]="未命中我方白名单：我方仅以席位i码出现，请人工确认我方产品及债券简称"
    return allrows

def write_rows_to_xlsx(rows, out_path):
    import openpyxl
    from openpyxl.styles import Font,PatternFill

    wb=openpyxl.Workbook()
    ws=wb.active
    ws.title="脚本提取"
    ws.append(OUTPUT_COLUMNS)
    for c in ws[1]:
        c.font=Font(bold=True,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor="305496")
    for r in rows:
        ws.append([r.get(c,"") for c in OUTPUT_COLUMNS])
    wb.save(out_path)

def main():
    base=os.path.dirname(os.path.abspath(__file__))
    files=sorted(glob.glob(os.path.join(base,"代投代缴-*.txt")))
    allrows=[]
    for f in files:
        txt=open(f,encoding='utf-8').read()
        allrows.extend(parse_text(txt))
    out=os.path.join(base,"现券预录单要素_脚本输出.xlsx")
    write_rows_to_xlsx(allrows, out)
    print("行数:",len(allrows),"->",out)
    # 命中白名单统计
    miss=sum(1 for r in allrows if r["我方账户"].startswith("(未"))
    print("未命中白名单行:",miss)

if __name__=="__main__":
    main()
