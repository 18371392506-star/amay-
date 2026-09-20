# coding: utf-8
"""锋铭出口单证：分别读取发票与装箱单，再填充企业专用模板。"""
from copy import copy, deepcopy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
import math
import re
import zipfile

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment


STATIC_DIR = Path(__file__).resolve().parent / "static" / "fengming"
COMPANY = "东莞市锋铭实业有限公司"
CUSTOMS_CODE = "44199639DT"
CREDIT_CODE = "91441900058561222U"
DEFAULT_BUYER = "香港锋铭实业有限公司"
DEFAULT_BUYER_ADDRESS = "香港九龍旺角花園街2-16號好景商業中心28樓05室"
DEFAULT_CONSIGNEE = "SMR Automotive Mirror Technology Hungary BT"
MONEY_TOLERANCE = Decimal("0.01")

# 以下是用户提供并确认的企业规则，不作自动商品归类推断。
PRODUCTS = {
    "mould": {
        "name": "注塑模具/用于生产汽车后视镜配件", "code": "8480719090", "unit": "套",
        "description": "其他塑料或橡胶用注模",
        "elements": "1.品牌类型：无品牌；2.出口享惠：不享惠；3.产品用途：用于生产汽车后视镜配件；"
        "4.适用材料：塑胶；5.品牌：无牌；6.型号：{model}；"
        "7.原理：塑胶颗粒经高温溶解后注入模具中，经冷却成型；8.材质：钢材",
    },
    "fixture": {
        "name": "夹具", "code": "8466200090", "unit": "个",
        "elements": "1.品牌类型：无品牌；2.出口享惠：不享惠；"
        "3.产品用途：用于固定塑胶产品，以方便量测；4.品牌：无牌；5.型号：{model}",
    },
    "plastic": {
        "name": "汽车后视镜塑胶件", "code": "8708299000", "unit": "个",
        "elements": "1.品牌类型：无品牌；2.出口享惠：不享惠；3.适用车型：通用型；"
        "4.非成套散件；5.品牌：无牌；6.型号：{model}",
    },
}


def _text(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _key(value):
    return re.sub(r"[\s:：()（）._/\\-]+", "", _text(value)).upper()


def _number(value, label):
    text = _text(value).replace(",", "").replace("，", "")
    matches = re.findall(r"[-+]?\d+(?:\.\d+)?", text)
    if len(matches) != 1:
        raise ValueError(f"{label}不是有效数字：{text or '空白'}")
    try:
        result = Decimal(matches[0])
    except InvalidOperation as exc:
        raise ValueError(f"{label}不是有效数字：{text}") from exc
    if not result.is_finite() or result < 0:
        raise ValueError(f"{label}不能为负数或无效值。")
    return result


def _fmt(value):
    return format(value, "f").rstrip("0").rstrip(".") if "." in format(value, "f") else str(value)


def _date(value):
    if isinstance(value, (date, datetime)):
        return value.date() if isinstance(value, datetime) else value
    value = _text(value)
    match = re.search(r"(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})", value)
    if not match:
        raise ValueError(f"无法识别日期：{value or '空白'}。请使用年-月-日格式。")
    try:
        return date(*(int(x) for x in match.groups()))
    except ValueError as exc:
        raise ValueError(f"日期无效：{value}") from exc


def _sheet(book, kind):
    candidates = []
    for name in book.sheet_names:
        key = _key(name)
        if (kind == "invoice" and ("发票" in key or key == "INVOICE")) or (
            kind == "packing" and ("装箱单" in key or key in {"PACKINGLIST", "PACKING"})):
            candidates.append(name)
    if len(candidates) != 1:
        label = "发票 INVOICE" if kind == "invoice" else "装箱单 PACKING LIST"
        raise ValueError(f"请提供唯一的【{label}】工作表，目前找到 {len(candidates)} 个。")
    return book.parse(candidates[0], header=None, dtype=object).fillna("")


def _header(frame, kind):
    aliases = {
        "name": ("中文品名", "品名", "产品名称", "DESCRIPTION"),
        "contract": ("合同号", "合同编号", "CONTRACTNO"),
        "model": ("型号", "MODEL", "PARTNO"),
        "brand": ("牌子", "品牌", "BRAND"),
        "qty": ("数量", "QUANTITY", "QTY"),
        "price": ("单价", "UNITPRICE"),
        "amount": ("总价", "金额", "AMOUNT", "TOTALAMOUNT"),
        "nw": ("总净重", "净重", "NETWEIGHT", "NW"),
        "gw": ("总毛重", "毛重", "GROSSWEIGHT", "GW"),
        "packages": ("箱量", "箱数", "包装件数", "CTNS"),
    }
    required = {"name", "contract", "model", "qty"}
    required |= {"price", "amount"} if kind == "invoice" else {"nw", "gw", "packages"}
    for index, row in frame.iterrows():
        mapping = {}
        for col, value in enumerate(row):
            key = _key(value)
            for field, names in aliases.items():
                if any(key.startswith(alias) for alias in names):
                    mapping.setdefault(field, col)
                    break
        if required <= mapping.keys():
            return index, mapping
    raise ValueError(f"{'发票' if kind == 'invoice' else '装箱单'}缺少必要表头，请检查品名、合同号、型号、数量及金额/重量列。")


def _metadata(frame, end, labels):
    for _, row in frame.iloc[:end].iterrows():
        for col, raw in enumerate(row):
            value = _text(raw)
            for label in labels:
                match = re.match(r"\s*" + re.escape(label) + r"\s*[:：]\s*(.*)", value, re.I)
                if not match:
                    continue
                if match.group(1).strip():
                    return match.group(1).strip()
                for adjacent in row.iloc[col + 1:]:
                    if _text(adjacent):
                        return adjacent
    return ""


def _category(name):
    key = _key(name)
    if key in {"注塑模具", "注塑模具用于生产汽车后视镜塑胶件", "注塑模具用于生产汽车后视镜配件"}:
        return "mould"
    if key == "夹具":
        return "fixture"
    if key == "汽车后视镜塑胶件":
        return "plastic"
    raise ValueError(f"尚未配置商品【{name}】的申报规则，请先补充品名、编码和申报要素。")


def _total_row(row, name_col):
    if _text(row.iloc[name_col]):
        return bool(re.fullmatch(r"(?:总计|合计|TOTAL)[:：]?", _text(row.iloc[name_col]), re.I))
    return any(re.fullmatch(r"(?:总计|合计|TOTAL)[:：]?", _text(v), re.I) for v in row)


def _required(value, label):
    value = _text(value)
    if not value:
        raise ValueError(f"缺少{label}，请检查上传文件或补充填写。")
    return value


def read_invoice_data(source):
    """支持 .xls/.xlsx；不使用样例值补齐缺失的交易数据。"""
    with pd.ExcelFile(source) as book:
        invoice, packing = _sheet(book, "invoice"), _sheet(book, "packing")
    inv_head, inv_cols = _header(invoice, "invoice")
    pack_head, pack_cols = _header(packing, "packing")
    for frame, header, label in ((invoice, inv_head, "发票"), (packing, pack_head, "装箱单")):
        if not any(_key(value) == _key(COMPANY) for row in frame.iloc[:header].values for value in row):
            raise ValueError(f"{label}抬头未找到【{COMPANY}】，请确认选择了正确企业。")
    invoice_date = _date(_metadata(invoice, inv_head, ["日期", "Date"]))
    destination = _required(_metadata(invoice, inv_head, ["目的地", "目的国", "Destination"]), "发票目的地")
    pack_destination = _text(_metadata(packing, pack_head, ["目的地", "目的国", "Destination"]))
    if pack_destination and _key(destination) != _key(pack_destination):
        raise ValueError("发票与装箱单目的地不一致，请核对。")
    incoterms = _key(_metadata(invoice, inv_head, ["成交方式", "Incoterms"]))
    if incoterms == "CFR":
        incoterms = "C&F"
    if incoterms not in {"C&F", "CIF", "FOB", "EXW"}:
        raise ValueError("未识别到成交方式，请在发票中填写 C&F、CFR、CIF、FOB 或 EXW。")
    price_header = _key(invoice.iloc[inv_head, inv_cols["price"]])
    amount_header = _key(invoice.iloc[inv_head, inv_cols["amount"]])
    currencies = {c for c in ("EUR", "USD", "CNY", "HKD") if c in price_header + amount_header}
    if len(currencies) != 1:
        raise ValueError("请在单价/总价表头中注明一致的币种 EUR、USD、CNY 或 HKD。")
    currency = currencies.pop()
    items, contracts, invoice_totals = [], set(), []
    for index, row in invoice.iloc[inv_head + 1:].iterrows():
        get = lambda field: row.iloc[inv_cols[field]]
        name = _text(get("name"))
        summary = (not name and not any(_text(get(f)) for f in ("model", "contract", "price"))
                   and any(_text(get(f)) for f in ("amount", "qty")))
        if _total_row(row, inv_cols["name"]) or summary:
            invoice_totals.append(_number(get("amount"), "发票总金额"))
            continue
        if not name:
            if any(_text(get(f)) for f in ("qty", "price", "model")):
                raise ValueError(f"发票第 {index + 1} 行缺少品名，不能忽略有数据的商品行。")
            continue
        category = _category(name)
        rule = PRODUCTS[category]
        model = _required(get("model"), f"发票第 {index + 1} 行型号")
        contract = _required(get("contract"), f"发票第 {index + 1} 行合同号")
        brand = _text(get("brand")) if "brand" in inv_cols else ""
        if _key(brand) not in {"无牌", "无品牌", "NOBRAND"}:
            raise ValueError(f"商品【{name}】品牌为【{brand or '空白'}】，当前申报规则仅适用于无品牌商品。")
        qty = _number(get("qty"), f"{name}数量")
        price = _number(get("price"), f"{name}单价")
        amount = _number(get("amount"), f"{name}金额")
        if qty <= 0 or qty != qty.to_integral_value():
            raise ValueError(f"商品【{name}】数量必须为正整数。")
        if abs(qty * price - amount) > MONEY_TOLERANCE:
            raise ValueError(f"商品【{name}】数量×单价与金额不一致，请核对发票。")
        items.append(dict(source_name=name, name=rule["name"], category=category,
                          model=model, qty=qty, price=price, amount=amount,
                          code=rule["code"], unit=rule["unit"]))
        contracts.add(contract)
    if not items:
        raise ValueError("发票未找到有效商品。")
    if len(contracts) != 1:
        raise ValueError("发票包含多个合同号，请按合同拆分后生成。")
    contract_number = contracts.pop()
    total_amount = sum((item["amount"] for item in items), Decimal(0))
    if len(invoice_totals) != 1 or abs(invoice_totals[0] - total_amount) > MONEY_TOLERANCE:
        raise ValueError("发票总金额缺失、重复或与商品金额合计不一致，请核对。")

    packing_totals, packing_items, item_weights, gross_weights = [], {}, [], []
    for index, row in packing.iloc[pack_head + 1:].iterrows():
        get = lambda field: row.iloc[pack_cols[field]]
        if _total_row(row, pack_cols["name"]):
            packing_totals.append(row)
            continue
        name = _text(get("name"))
        if not name:
            if any(_text(get(f)) for f in ("qty", "model", "nw", "gw", "packages")):
                raise ValueError(f"装箱单第 {index + 1} 行缺少品名，请检查合并单元格或商品数据。")
            continue
        category = _category(name)
        model = _required(get("model"), f"装箱单第 {index + 1} 行型号")
        if _text(get("contract")) != contract_number:
            raise ValueError("发票与装箱单合同号不一致。")
        key = (category, model)
        qty = _number(get("qty"), f"装箱单{name}数量")
        packing_items[key] = packing_items.get(key, Decimal(0)) + qty
        item_weights.append(_number(get("nw"), f"装箱单{name}净重"))
        # 共箱时毛重只填写在合并区域首行，空白行不重复计重。
        if _text(get("gw")):
            gross_weights.append(_number(get("gw"), f"装箱单{name}毛重"))
    invoice_items = {}
    for item in items:
        key = (item["category"], item["model"])
        invoice_items[key] = invoice_items.get(key, Decimal(0)) + item["qty"]
    if invoice_items != packing_items:
        raise ValueError("发票与装箱单的商品、型号或数量不一致，请核对。")
    if len(packing_totals) != 1:
        raise ValueError("装箱单需包含唯一的 TOTAL/总计/合计行。")
    total = packing_totals[0]
    net_weight = _number(total.iloc[pack_cols["nw"]], "总净重")
    gross_weight = _number(total.iloc[pack_cols["gw"]], "总毛重")
    packages_text = _text(total.iloc[pack_cols["packages"]])
    package_match = re.fullmatch(r"\s*(\d+)\s*(?:箱|CTNS?)\s*[/／]\s*(.+?)\s*", packages_text, re.I)
    if not package_match:
        raise ValueError("装箱单总箱数/包装应类似“2箱/胶合板”，请核对汇总行。")
    total_packages = int(package_match.group(1))
    pack_type = package_match.group(2)
    if pack_type == "胶合板":
        pack_type = "胶合板箱"
    if total_packages <= 0 or net_weight <= 0 or gross_weight < net_weight:
        raise ValueError("箱数必须大于0，净重必须大于0，毛重不能小于净重。")
    if abs(sum(item_weights) - net_weight) > Decimal("0.01"):
        raise ValueError("装箱单商品净重合计与总净重不一致。")
    if abs(sum(gross_weights) - gross_weight) > Decimal("0.01"):
        raise ValueError("装箱单毛重合计与总毛重不一致，请检查共箱合并区域。")
    total_qty = sum((item["qty"] for item in items), Decimal(0))
    if _number(total.iloc[pack_cols["qty"]], "装箱单总数量") != total_qty:
        raise ValueError("装箱单总数量与商品数量合计不一致。")
    return dict(company_name=COMPANY, date=invoice_date, destination=destination,
                contract_number=contract_number, currency=currency, incoterms=incoterms,
                items=items, total_amount=total_amount, total_packages=total_packages,
                pack_type=pack_type, net_weight=net_weight, gross_weight=gross_weight)


def _font(run, size=10):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")


def _paragraphs(doc):
    yield from doc.paragraphs
    for table in doc.tables:
        seen = set()
        for row in table.rows:
            for cell in row.cells:
                if cell._tc not in seen:
                    seen.add(cell._tc)
                    yield from cell.paragraphs


def _replace(doc, values):
    # 在跨 run 的占位符上仅替换对应范围，保留其他文本和格式。
    for paragraph in _paragraphs(doc):
        for token, value in values.items():
            while token in paragraph.text:
                start = paragraph.text.index(token)
                end = start + len(token)
                offset, first = 0, True
                for run in paragraph.runs:
                    old = run.text
                    left, right = offset, offset + len(old)
                    offset = right
                    if right <= start or left >= end:
                        continue
                    prefix = old[:max(0, start - left)]
                    suffix = old[max(0, end - left):]
                    run.text = prefix + (str(value) if first else "") + suffix
                    first = False
                if first:
                    raise ValueError(f"无法替换模板字段 {token}")
    remaining = [p.text for p in _paragraphs(doc) if re.search(r"\{\{[A-Z_]+\}\}", p.text)]
    if remaining:
        raise ValueError("文档模板存在未填写字段。")


def create_declaration_elements(data):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10.5)
    normal.font.name = "宋体"
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("商品申报要素清单")
    _font(r, 16)
    r.bold = True
    for index, item in enumerate(data["items"], 1):
        rule = PRODUCTS[item["category"]]
        p = doc.add_paragraph(f"{index}、商品编码：{item['code']}  {item['name']}")
        p.paragraph_format.keep_with_next = True
        p.runs[0].bold = True
        if rule.get("description"):
            p = doc.add_paragraph("商品描述：" + rule["description"])
            p.paragraph_format.keep_with_next = True
        doc.add_paragraph("申报要素：" + rule["elements"].format(model=item["model"]))
    stream = BytesIO()
    doc.save(stream)
    return stream.getvalue()


def create_sales_contract(data, inputs):
    doc = Document(STATIC_DIR / "购销合同模板.docx")
    table = doc.tables[0]
    prototype = table.rows[3]._tr
    anchor = prototype
    for index, item in enumerate(data["items"], 1):
        node = deepcopy(prototype)
        anchor.addnext(node)
        anchor = node
        row = table.rows[3 + index]
        values = {0: str(index), 2: item["name"], 3: f"{_fmt(item['qty'])}{item['unit']}",
                  4: f"{data['currency']}\n{_fmt(item['price'])}",
                  5: f"{data['currency']}\n{item['amount']:.2f}",
                  7: f"成交方式:{data['incoterms']}" if index == len(data["items"]) else ""}
        for col, value in values.items():
            cell = row.cells[col]
            cell.text = value
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.space_before = Pt(3)
                for run in p.runs:
                    _font(run, 9)
        row.height = None
        cant_split = OxmlElement("w:cantSplit")
        row._tr.get_or_add_trPr().append(cant_split)
    prototype.getparent().remove(prototype)
    contract_date = _date(inputs["contract_date"])
    _replace(doc, {
        "{{BUYER}}": inputs["buyer_name"], "{{BUYER_ADDRESS}}": inputs.get("buyer_address", ""),
        "{{BUYER_PHONE}}": inputs.get("buyer_phone", ""),
        "{{CONTRACT}}": data["contract_number"],
        "{{CONTRACT_DATE}}": f"{contract_date.year}年{contract_date.month}月{contract_date.day}日",
        "{{DESTINATION}}": data["destination"],
        "{{TOTAL}}": f"{data['currency']} {data['total_amount']:.2f}",
    })
    stream = BytesIO()
    doc.save(stream)
    return stream.getvalue()


def create_export_declaration(data, inputs):
    wb = load_workbook(STATIC_DIR / "出口报关单模板.xlsx")
    ws = wb.active
    # 模板第11行为明细原型，第12、13行为完整页脚，按商品条数顺移。
    count = len(data["items"])
    extra = count - 1
    footer_values = []
    for row in (12, 13):
        for col in range(1, 12):
            c = ws.cell(row, col)
            footer_values.append((row, col, c.value, copy(c._style), copy(c.alignment)))
    footer_heights = {r: ws.row_dimensions[r].height for r in (12, 13)}
    footer_merges = [str(m) for m in ws.merged_cells.ranges if m.min_row >= 12]
    for merged in footer_merges:
        ws.unmerge_cells(merged)
    if extra:
        ws.insert_rows(12, extra)
    for row, col, value, style, alignment in footer_values:
        cell = ws.cell(row + extra, col, value)
        cell._style, cell.alignment = style, alignment
    for row, height in footer_heights.items():
        ws.row_dimensions[row + extra].height = height
    for merged in footer_merges:
        from openpyxl.worksheet.cell_range import CellRange
        shifted = CellRange(merged)
        shifted.shift(row_shift=extra)
        ws.merge_cells(str(shifted))
    updates = {
        "A3": f"境内发货人:{CUSTOMS_CODE} {COMPANY}({CREDIT_CODE})",
        "A4": "境外收货人：\n" + inputs["consignee"],
        "A5": f"生产销售单位\n{CUSTOMS_CODE} {COMPANY}({CREDIT_CODE})",
        "A6": "合同协议号\n" + data["contract_number"],
        "D6": "贸易国（地区）\n" + inputs["trade_country"],
        "F6": "运抵国（地区）\n" + data["destination"],
        "A7": "包装种类:" + inputs.get("pack_type", data["pack_type"]),
        "D7": f"件数:{data['total_packages']}件",
        "E7": f"毛重（千克):{_fmt(data['gross_weight'])}",
        "F7": f"净重（千克):{_fmt(data['net_weight'])}",
        "G7": "成交方式:" + data["incoterms"],
        "H7": "运费：" + inputs.get("freight", ""),
        "I7": "保费：" + inputs.get("insurance", ""),
        "J7": "杂费：" + inputs.get("other_fees", ""),
    }
    for address, value in updates.items():
        ws[address] = value
        ws[address].alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    currencies = {"EUR": "欧元", "USD": "美元", "CNY": "人民币", "HKD": "港币"}
    styles = [copy(ws.cell(11, col)._style) for col in range(1, 12)]
    for index, item in enumerate(data["items"]):
        row = 11 + index
        # 报关单字段为“商品名称及规格型号”，同时写入型号以便核对。
        name = f"{item['name']}\n型号：{item['model']}"
        values = [index + 1, item["code"], name, f"{_fmt(item['qty'])}{item['unit']}",
                  float(item["price"]), float(item["amount"]), currencies[data["currency"]],
                  "中国", data["destination"], "东莞", "照章征税"]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell._style = copy(styles[col - 1])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if col in (5, 6):
                cell.number_format = "0.00"
        name_lines = math.ceil(len(item["name"]) / 10) + math.ceil((len(item["model"]) + 3) / 16)
        ws.row_dimensions[row].height = max(54, 15 * name_lines)
    ws.print_area = f"A1:K{13 + extra}"
    ws.print_title_rows = "1:10"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if count <= 5 else 0
    stream = BytesIO()
    wb.save(stream)
    wb.close()
    return stream.getvalue()


def generate_documents(data, inputs):
    for field, label in (("buyer_name", "合同买方"), ("consignee", "境外收货人"),
                         ("trade_country", "贸易国"), ("contract_date", "合同日期")):
        _required(inputs.get(field), label)
    stamp = data["date"].strftime("%Y%m%d")
    documents = {
        f"锋铭_申报要素_{stamp}.docx": create_declaration_elements(data),
        f"锋铭_购销合同_{stamp}.docx": create_sales_contract(data, inputs),
        f"锋铭_出口报关单_{stamp}.xlsx": create_export_declaration(data, inputs),
    }
    # ZIP只包含生成结果；使用内存缓冲区，避免并发请求覆盖临时文件。
    archive = BytesIO()
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in documents.items():
            z.writestr(name, content)
    return documents, archive.getvalue()


def render():
    import hashlib
    import streamlit as st

    st.header("东莞锋铭 · 出口单证自动生成")
    st.caption("上传含发票和装箱单的 Excel，生成申报要素、购销合同和出口报关单。")
    uploaded = st.file_uploader("1. 上传发票及装箱单 (.xls / .xlsx)", type=["xls", "xlsx"], key="fm_upload")
    if uploaded is None:
        for key in ("fm_result", "fm_source", "fm_data"):
            st.session_state.pop(key, None)
        return
    content = uploaded.getvalue()
    fingerprint = hashlib.sha256(content).hexdigest()
    if fingerprint != st.session_state.get("fm_source"):
        st.session_state.pop("fm_result", None)
        st.session_state.pop("fm_data", None)
        st.session_state.pop("fm_source", None)
        st.session_state.pop("fm_contract_date", None)
        try:
            st.session_state["fm_data"] = read_invoice_data(BytesIO(content))
        except Exception as exc:
            st.error(str(exc))
            return
        st.session_state["fm_source"] = fingerprint
    data = st.session_state["fm_data"]
    st.success(f"读取成功：合同 {data['contract_number']}，{len(data['items'])} 项商品，"
               f"合计 {data['total_amount']:,.2f} {data['currency']}。")
    st.write(f"发票日期：{data['date']}　目的地：{data['destination']}　成交方式：{data['incoterms']}")
    st.write(f"包装：{data['total_packages']}箱（{data['pack_type']}）　"
             f"净重：{_fmt(data['net_weight'])} kg　毛重：{_fmt(data['gross_weight'])} kg")
    st.dataframe([{"品名": i["name"], "型号": i["model"], "数量": f"{_fmt(i['qty'])}{i['unit']}",
                   "单价": _fmt(i["price"]), "金额": f"{i['amount']:.2f}"} for i in data["items"]],
                 hide_index=True, use_container_width=True)
    st.markdown("**2. 补充单证信息**")
    c1, c2 = st.columns(2)
    buyer = c1.text_input("合同买方", DEFAULT_BUYER, key="fm_buyer")
    consignee = c2.text_input("境外收货人", DEFAULT_CONSIGNEE, key="fm_consignee")
    contract_date = c1.date_input("合同日期（单独填写）", value=None, key="fm_contract_date",
                                  help="合同日期可与发票日期不同，请按本次合同填写。")
    trade_country = c2.text_input("贸易国（合同买方所在国家/地区）", "中国香港", key="fm_trade_country")
    buyer_address = st.text_input("买方地址", DEFAULT_BUYER_ADDRESS, key="fm_buyer_address")
    buyer_phone = st.text_input("买方电话", "00852-39622458", key="fm_buyer_phone")
    c3, c4, c5 = st.columns(3)
    freight = c3.text_input("运费", key="fm_freight")
    insurance = c4.text_input("保费", key="fm_insurance")
    other_fees = c5.text_input("杂费", key="fm_other_fees")
    inputs = dict(buyer_name=buyer.strip(), consignee=consignee.strip(), contract_date=contract_date,
                  trade_country=trade_country.strip(), buyer_address=buyer_address.strip(),
                  buyer_phone=buyer_phone.strip(), freight=freight.strip(), insurance=insurance.strip(),
                  other_fees=other_fees.strip(), pack_type=data["pack_type"])
    signature = (fingerprint, tuple((k, str(v)) for k, v in inputs.items()))
    previous = st.session_state.get("fm_result")
    if previous and previous[0] != signature:
        st.session_state.pop("fm_result", None)
    if st.button("生成锋铭单证", type="primary", key="fm_generate"):
        st.session_state.pop("fm_result", None)
        try:
            with st.spinner("正在生成三份单证…"):
                documents, archive = generate_documents(data, inputs)
            st.session_state["fm_result"] = (signature, documents, archive)
        except Exception as exc:
            st.error(f"生成失败：{exc}")
    result = st.session_state.get("fm_result")
    if result:
        st.download_button("下载锋铭单证 ZIP", result[2],
                           f"锋铭_单证_{data['date']:%Y%m%d}.zip", "application/zip", key="fm_zip")
        for index, (name, value) in enumerate(result[1].items()):
            mime = ("application/vnd.openxmlformats-officedocument.wordprocessingml.document" if name.endswith(".docx")
                    else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button(name, value, name, mime, key=f"fm_download_{index}")
