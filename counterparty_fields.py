# -*- coding: utf-8 -*-
"""对方账户与过券机构解析。

账户按字段边界提取；“号/计划/年金”等只用于判断账户类型，不再作为截断位置。
"""
import re


_FULLWIDTH_TRANS = str.maketrans({chr(0xFF01 + i): chr(0x21 + i) for i in range(0x7E - 0x21 + 1)})
MINE_PREFIXES = ("中信信托", "粤财信托", "财信信托")
ACCOUNT_HINT_WORDS = ("号", "计划", "组合", "基金", "年金", "资管", "产品", "私募")

RE_ORG = re.compile(
    r'([一-龥]{2,20}(?:证券资管|证券自营|证券机构经纪|证券|资管|基金|信托|银行|期货|养老|保险|人寿|投顾|财富|资产))'
)
RE_DEALER = re.compile(r'交易商(?:代码|号)?\s*[:：]?\s*(\d{6})')
RE_TRADER = re.compile(r'交易员(?:代码|号)?\s*[:：]?\s*([0-9A-Z]{8})(?![0-9A-Z])', re.I)
RE_TRADER2 = re.compile(
    r'交易员(?:及交易员代码|代码|名称|号)?\s*[：:]?\s*[一-龥]{2,4}\s*[（(]?\s*([0-9A-Z]{7,8})(?![0-9A-Z])',
    re.I,
)
RE_BOND = re.compile(r'(?:(\d{6})\s*[.\-/]?\s*(SH|SZ)|(SH|SZ)\s*[:.\-/]?\s*(\d{6}))', re.I)
RE_YD = re.compile(r'约(?:定号)?[：:\s]*([0-9]+(?:\s*[+＋、，,]\s*[0-9]+)*)')

_PROD_TAIL = (
    r'(?:\d{1,3}个月)?'
    r'(?:集合资产管理计划|资产管理计划|职业年金计划|企业年金计划|年金计划)?'
)
RE_PROD = re.compile(
    r'([一-龥A-Za-z0-9－\-—（）()·]{2,120}?'
    r'(?:\d+M?\d*号' + _PROD_TAIL +
    r'|集合资产管理计划|资产管理计划|职业年金计划|企业年金计划|企业年金|年金计划|组合|基金|私募基金))'
)
RE_LEAD_ENUM = re.compile(r'^[ \t]*(?:\d{1,2}\s*[.、）)．]|[①②③④⑤⑥⑦⑧⑨⑩]|[a-zA-Z]\s*[.、)])\s*')

_ROLE_FIELD = (
    r'(?:交易日期|清算速度|结算速度|结算方式|交易方向|方向|债券代码|债券简称|券简称|债券名称|'
    r'到期收益率|行权收益率|收益率|YTM|原始净价|交易净价|成交净价|净价|交易规模|成交规模|规模|'
    r'我方账户|本方账户|我司账户|我方产品|本方产品|对方账户|对手方账户|交易对手账户|对方产品|'
    r'交易商代码|交易员代码|交易主体代码|约定号|备注)'
)
RE_MINE_ACCOUNT_LABEL = re.compile(
    r'(?:我方账户|本方账户|我司账户|我方产品|本方产品)\s*[:：]\s*[“"\'‘]?\s*'
    r'(.+?)(?=\s*(?:[，,；;]\s*)?(?:' + _ROLE_FIELD + r'\s*[:：]|\n|$))'
)
RE_COUNTER_ACCOUNT_LABEL = re.compile(
    r'(?:对方账户|对手方账户|交易对手账户|对方产品)\s*[:：]\s*[“"\'‘]?\s*'
    r'(.+?)(?=\s*(?:[，,；;]\s*)?(?:' + _ROLE_FIELD + r'\s*[:：]|\n|$))'
)

DEALER_NAME = {
    "000262": "中信证券", "000038": "中信证券华南", "000032": "广发证券", "000680": "中信建投证券",
    "000039": "世纪证券", "000613": "首创证券", "006206": "国泰基金", "000028": "方正证券",
    "000287": "东吴证券", "000664": "财通证券", "006285": "国联基金", "006205": "富国基金",
    "006281": "太平基金", "000128": "兴业证券", "000001": "国信证券", "000612": "国泰海通",
    "000058": "华鑫证券", "000695": "申港证券", "000316": "第一创业", "007101": "华泰资产",
    "007130": "英大资产", "000402": "东方财富", "000025": "光大证券", "006223": "宝惠",
    "000657": "中邮证券", "007128": "东方证券资管",
}
DEALER_ORG_NAMES = tuple(sorted(set(DEALER_NAME.values()), key=len, reverse=True))

_INVALID_ACCOUNT_EXACT = {
    "我方", "本方", "我司", "对方", "对手方", "买方", "卖方", "发单", "先发",
    "买家先发", "卖家先发", "买方发单", "卖方发单", "交易", "交易信息",
    "要素", "要素已定", "账户", "产品", "机构经纪", "(点)", "(发)",
}


def _norm(text):
    text = (text or "").translate(_FULLWIDTH_TRANS)
    text = text.replace("【", " ").replace("】", " ").replace("（", "(").replace("）", ")")
    return re.sub(r'[\u3000\xa0]+', ' ', text)


def _compact(text):
    return re.sub(r'\s+', '', _norm(text))


def looks_like_account_name(text):
    value = _compact(text)
    if not value:
        return False
    if any(word in value for word in ("计划", "组合", "基金", "年金", "私募")):
        return True
    if re.search(r'\d+(?:个)?月', value) or any(value.startswith(prefix) for prefix in MINE_PREFIXES):
        return True
    return bool(re.fullmatch(r'[一-龥]{2,20}\d+M?\d*号', value))


def looks_like_plain_org(text):
    value = _compact(text)
    m = RE_ORG.search(value) if value else None
    return bool(m and m.group(1) == value)


def strip_custodian(text):
    """仅剥离年金账户后面的银行托管机构，保留产品名称内部的短横线。"""
    value = _norm(text).strip()
    parts = re.split(r'\s*[－—-]\s*', value)
    if len(parts) < 2:
        return value
    left, right = '－'.join(parts[:-1]).strip(), parts[-1].strip()
    if re.search(r'(?:职业年金计划|企业年金计划|企业年金|年金计划)$', left) and "银行" in right:
        return left
    return value


def extract_full_account_from_line(line):
    """按主体代码/规模/约定号边界读取完整账户，保留非数值括号。"""
    text = RE_LEAD_ENUM.sub('', _norm(line or '')).strip()
    if not text:
        return ""
    text = re.sub(
        r'^(?:交易主体全称|主体全称|交易主体名称|主体名称|交易主体简称|主体简称|对应账户)\s*[:：]?\s*',
        '',
        text,
    )
    text = re.sub(r'^(?:交易主体代码\s*[:：]?\s*)?36\d{8}\s*', '', text)
    text = re.split(r'\s*约(?:定号)?\s*[:：]?\s*\d', text, maxsplit=1)[0]
    text = re.sub(r'[（(]\s*\d+(?:\.\d+)?\s*(?:万|[wW]|亿|[eE])\s*[)）]', ' ', text)
    text = re.sub(r'\s+\d+(?:\.\d+)?\s*(?:万|[wW]|亿|[eE])\s*$', '', text)
    text = re.sub(r'\s+\d+(?:\.\d+)?\s*$', '', text)
    text = re.sub(r'\s+(?:交易商|交易员|交易主体代码)\s*[:：].*$', '', text)
    text = strip_custodian(text.strip(' \t：:，,；;'))
    return re.sub(r'\s+', '', text)


def extract_labeled_account(seg):
    match = re.search(
        r'账户\s*[：:]?\s*[“"\'‘]?\s*(.+?)\s*'
        r'(?=(?:本方交易商简称|交易主体(?:代码|名称|简称|全称)|交易商(?:代码|号)?|交易员(?:代码|号)?|约(?:定号)?|\)|）|\n|$))',
        seg or "",
    )
    return _norm(match.group(1)).strip(' \t：:，,；;“”"\'‘’（）()') if match else ""


def extract_role_account(seg, role, with_pos=False):
    pattern = RE_MINE_ACCOUNT_LABEL if role == "mine" else RE_COUNTER_ACCOUNT_LABEL
    match = pattern.search(seg or "")
    if not match:
        return ("", -1) if with_pos else ""
    account = re.sub(r'\s+', '', _norm(match.group(1)).strip(' \t：:，,；;“”"\'‘’（）()'))
    return (account, match.start()) if with_pos else account


def pick_org(seg):
    match = RE_ORG.search(seg or "")
    return match.group(1) if match else ""


def pick_prod(seg, mine_related=lambda _: False):
    labeled = extract_labeled_account(seg)
    if labeled and not mine_related(labeled):
        return labeled
    for line in (seg or "").splitlines():
        if re.match(r'\s*(?:交易主体(?:全称|名称|简称)|主体(?:全称|名称|简称)|(?:交易主体代码\s*[:：]?\s*)?36\d{8})', line):
            account = extract_full_account_from_line(line)
            if account and looks_like_account_name(account) and not mine_related(account):
                return account
    for match in RE_PROD.finditer(seg or ""):
        if not mine_related(match.group(1)):
            return match.group(1)
    return ""


def extract_subject_name(seg):
    for line in (seg or "").splitlines():
        if re.match(r'\s*(?:交易主体全称|主体全称|交易主体名称|主体名称|交易主体简称|主体简称)\s*[:：]?', line):
            account = extract_full_account_from_line(line)
            if account:
                return account
        match = re.match(r'\s*交易(?:商)?主体(?!代码|名称|简称|全称)\s*[:：]?\s*(.*)$', line)
        if match and match.group(1).strip():
            account = extract_full_account_from_line(match.group(1))
            if account:
                return account
    match = re.search(r'交易主体代码[（(]([一-龥A-Za-z0-9]{2,40})[)）]', seg or "")
    return extract_full_account_from_line(match.group(1)) if match else ""


def split_org_from_account(account):
    value = (account or "").strip()
    match = re.match(r'^([一-龥]{2,20}(?:证券资管|证券自营|证券机构经纪|证券|资管|基金|信托|银行|期货|保险|养老|人寿|财富|资产))(.+)$', value)
    if not match:
        return "", value
    org, tail = match.group(1), match.group(2).strip()
    if not tail or looks_like_plain_org(value) or looks_like_account_name(tail) or any(word in tail for word in ACCOUNT_HINT_WORDS):
        return org, value
    return "", value


def normalize_visible_org(text):
    value = _norm(text).strip(' \t：:，,；;')
    if not value:
        return ""
    value = re.split(r'[，,；;（(]', value, maxsplit=1)[0].strip()
    value = re.sub(r'(?:证券机构经纪|机构经纪|自营)$', '', value).strip(' \t：:，,；;')
    match = RE_ORG.match(value)
    return match.group(1) if match else (value if looks_like_plain_org(value) else "")


def clean_counterparty_candidate(text):
    value = _norm(text).strip()
    if not value:
        return ""
    value = value.translate(str.maketrans('', '', '“”‘’"\''))
    value = re.split(
        r'交易商(?:代码|号)?|交易员(?:代码|号)?|交易主体(?:代码|名称|简称|全称)|本方交易商简称|'
        r'(?:我方|本方|我司|对方|对手方|交易对手)?账户\s*[：:]|清算速度|结算速度|结算方式|清算方式',
        value,
        maxsplit=1,
    )[0]
    value = re.sub(r'^(?:深交|上交|交易所)\s*[:：]\s*', '', value)
    value = re.sub(r'^(?:出给|to|from|买入方?|卖出方?|买方|卖方)\s*[:：]?\s*', '', value, flags=re.I)
    value = re.sub(r'i\d{9}|[A-Za-z]*[Zz]\d{5,}', ' ', value)
    value = RE_YD.sub(' ', value)
    value = re.sub(r'约(?:定号)?\s*$', ' ', value)
    value = re.sub(r'(?<![A-Za-z0-9.])\d{2,6}(?:\s*[+＋]\s*\d{2,6})+(?![A-Za-z0-9])', ' ', value)
    value = re.sub(r'(?<![A-Za-z0-9.])\d+(?:\.\d+)?\s*(?:亿|[eE]|[kK](?:\s*[wW])?|[wW万])(?![A-Za-z0-9])', ' ', value)
    value = re.sub(r'20\d{2}[./-]\d{1,2}[./-]\d{1,2}|(?<!\d)20\d{6}(?!\d)', ' ', value)
    value = re.sub(r'(?<!\d)\d{1,2}[./-]\d{1,2}(?!\d)|(?:买券|卖券|交易|净价|行权|到期)', ' ', value)
    value = re.sub(r'(?<![A-Za-z0-9一-龥.])\d{2,6}(?![A-Za-z0-9一-龥.])', ' ', value)
    return _compact(value).strip('：:，,；;')


def is_valid_counterparty_account(text):
    value = clean_counterparty_candidate(text)
    if not value or value in _INVALID_ACCOUNT_EXACT:
        return False
    if re.fullmatch(r'(?:我方|本方|我司|对方|对手方|买方|卖方)?(?:发单|先发)', value):
        return False
    if re.search(r'(?:交易商|交易员|主体)(?:代码|名称|简称)?$', value):
        return False
    return not value.startswith(("号集合", "集合资产", "资产管理计划", "约定号"))


def resolve_counterparty_account(labeled="", subject_name="", product="", head="", mine_related=lambda _: False):
    subject_is_account = looks_like_account_name(subject_name)
    subject_is_broker_detail = bool(re.search(r'(?:证券自营|机构经纪)$', _compact(subject_name)))
    same_entity = subject_name and head and (_compact(head) in _compact(subject_name)) and not subject_is_broker_detail
    ordered = [
        labeled,
        subject_name if subject_is_account or same_entity else "",
        product,
        head,
        subject_name if not subject_is_account and not same_entity else "",
    ]
    for candidate in ordered:
        cleaned = clean_counterparty_candidate(candidate)
        if is_valid_counterparty_account(cleaned) and not mine_related(cleaned):
            return cleaned
    return ""


def _clean_org_candidate(text, mine_related=lambda _: False):
    raw = _norm(text).strip(' \t：:，,；;（）()')
    if not raw:
        return ""
    if raw in DEALER_ORG_NAMES:
        return raw
    candidate = normalize_visible_org(raw)
    if not candidate:
        match = RE_ORG.match(raw)
        candidate = match.group(1) if match else ""
    candidate = _compact(candidate)
    if not candidate or mine_related(candidate) or re.search(r'\d', candidate):
        return ""
    if candidate.startswith(("号", "集合", "计划", "组合", "产品")) or candidate in ("集合资产", "资产管理", "机构经纪"):
        return ""
    return candidate


def org_from_account(account, mine_related=lambda _: False):
    value = clean_counterparty_candidate(account)
    if not value:
        return ""
    for org in DEALER_ORG_NAMES:
        if value.startswith(org):
            return org
    prefix, _ = split_org_from_account(value)
    if prefix:
        return _clean_org_candidate(prefix, mine_related)
    match = RE_ORG.match(value)
    if match:
        return _clean_org_candidate(match.group(1), mine_related)
    return _clean_org_candidate(value, mine_related) if looks_like_plain_org(value) else ""


def find_known_dealer_code(seg):
    for match in re.finditer(r'(?<!\d)(\d{6})(?!\d)', seg or ""):
        if match.group(1) in DEALER_NAME:
            return match.group(1)
    return ""


def dealer_name_from_code(code):
    return DEALER_NAME.get(code, "")


def _explicit_guoquan(seg, mine_related):
    match = re.search(
        r'(?:过券(?:机构)?|经纪商|通道机构|交易商名称|本方交易商简称)\s*[：:]\s*([^\n，,；;（）()]{2,40})',
        seg or "",
    )
    return _clean_org_candidate(match.group(1), mine_related) if match else ""


def _dealer_org_on_line(seg, code, mine_related):
    if not code:
        return ""
    for line in (seg or "").splitlines():
        if code in line and "交易商" in line:
            org = _clean_org_candidate(line[line.find(code) + len(code):], mine_related)
            if org:
                return org
    return ""


def _leading_org(seg, mine_related):
    for raw in (seg or "").splitlines():
        line = RE_LEAD_ENUM.sub('', raw).strip()
        if not line:
            continue
        line = re.sub(r'^(?:出给|to|from|买入方?|卖出方?|买方|卖方)\s*[:：]?\s*', '', line, flags=re.I)
        match = RE_ORG.match(line)
        if match:
            org = _clean_org_candidate(match.group(1), mine_related)
            if org:
                return org
    return ""


def extract_account_execution_orgs(seg, mine_related=lambda _: False):
    """用角色证据而非固定句式，区分对手方的账户机构与执行机构。

    显式标签、已知交易商身份、机构类型、产品语义和交易元数据都只是加分证据；
    i/Z码或任一单独位置均不是必要条件。证据不足时不强拆，交给原有回落链处理。
    """
    for raw in (seg or "").splitlines():
        line = RE_LEAD_ENUM.sub('', raw).strip()
        if not line:
            continue
        line = re.sub(r'^(?:出给|to|from|买入方?|卖出方?|买方|卖方)\s*[:：]?\s*', '', line, flags=re.I)
        entities = []
        for match in RE_ORG.finditer(line):
            org = _clean_org_candidate(match.group(1), mine_related)
            if org:
                left = line[max(0, match.start() - 18):match.start()]
                right = line[match.end():match.end() + 50]
                execution_score = 0
                account_score = 0
                if re.search(r'(?:过券(?:机构)?|经纪商|通道机构|交易商名称|交易商简称)\s*[:：]?\s*$', left):
                    execution_score += 100
                if org in DEALER_ORG_NAMES:
                    execution_score += 60
                if re.search(r'(?:证券资管|证券自营|证券机构经纪|证券)$', org):
                    execution_score += 40
                if re.search(r'交易商|交易员|交易主体|席位|i\d{9}|[A-Za-z]*[Zz]\d{5,}', right, re.I):
                    execution_score += 15

                if re.search(r'(?:账户|产品|计划|组合|基金)所属(?:机构)?\s*[:：]?\s*$', left):
                    account_score += 100
                if re.search(r'(?:信托|银行|基金|资管|养老|保险|人寿|财富|资产)$', org):
                    account_score += 40
                if re.search(r'号|计划|组合|基金|年金|产品|私募', right):
                    account_score += 30
                entities.append((org, execution_score, account_score))
        if len(entities) < 2:
            continue
        if len({item[0] for item in entities}) == 1:
            return entities[0][0], entities[0][0]
        execution_index = max(range(len(entities)), key=lambda i: (entities[i][1], i))
        account_indexes = [i for i in range(len(entities)) if i != execution_index]
        account_index = max(account_indexes, key=lambda i: (entities[i][2], -i))
        if entities[execution_index][1] >= 40 and entities[account_index][2] >= 30:
            return entities[account_index][0], entities[execution_index][0]
    return "", ""


def _unique_execution_org_for_product(seg, account, mine_related):
    """产品型对方账户之外只有一个高置信机构时，将该机构作为过券候选。

    不仅凭“位于账户后面”判断；候选机构还必须具备已知交易商身份、
    证券类机构后缀，或紧邻交易商/交易员/席位代码等执行证据。
    """
    if not looks_like_account_name(account):
        return ""
    candidates = []
    for match in RE_ORG.finditer(seg or ""):
        org = _clean_org_candidate(match.group(1), mine_related)
        if not org:
            continue
        right = (seg or "")[match.end():match.end() + 50]
        execution_score = 0
        if org in DEALER_ORG_NAMES:
            execution_score += 60
        if re.search(r'(?:证券资管|证券自营|证券机构经纪|证券)$', org):
            execution_score += 40
        if re.search(r'交易商|交易员|交易主体|席位|i\d{9}|[A-Za-z]*[Zz]\d{5,}', right, re.I):
            execution_score += 15
        if execution_score >= 40:
            candidates.append(org)
    unique = list(dict.fromkeys(candidates))
    return unique[0] if len(unique) == 1 else ""


def resolve_guoquan(seg, dealer_code="", subject_name="", head="", account="", execution_org="", mine_related=lambda _: False):
    candidates = [
        _explicit_guoquan(seg, mine_related),
        _clean_org_candidate(execution_org, mine_related),
        _unique_execution_org_for_product(seg, account, mine_related),
        _dealer_org_on_line(seg, dealer_code, mine_related),
        dealer_name_from_code(dealer_code),
        _leading_org(seg, mine_related),
    ]
    if head and looks_like_plain_org(head):
        candidates.append(_clean_org_candidate(head, mine_related))
    if subject_name and not looks_like_account_name(subject_name):
        candidates.append(_clean_org_candidate(subject_name, mine_related))
    candidates.append(org_from_account(account, mine_related))
    for candidate in candidates:
        cleaned = _clean_org_candidate(candidate, mine_related)
        if cleaned:
            return cleaned
    return ""


def extract_counterparty_head(seg, mine_related=lambda _: False):
    for raw in (seg or "").splitlines():
        line = RE_LEAD_ENUM.sub('', raw).strip()
        if not line:
            continue
        line = re.split(r'交易商(?:代码|号)?|交易员(?:代码|号)?|交易主体(?:代码|名称|简称|全称)', line, maxsplit=1)[0]
        if RE_DEALER.search(line) or RE_TRADER.search(line) or RE_TRADER2.search(line) or any(x in line for x in ("交易员", "交易商", "交易主体")):
            continue
        line = clean_counterparty_candidate(line)
        if not is_valid_counterparty_account(line) or mine_related(line):
            continue
        if RE_BOND.search(line) or re.search(r'净价|行权|到期', line):
            continue
        return line
    return ""
