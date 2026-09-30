# coding: utf-8
"""康铭：独立发票、装箱单上传，全局享惠状态、独立报关单生成。"""
from io import BytesIO
from decimal import Decimal
from datetime import datetime
from zoneinfo import ZoneInfo
import re
import zipfile
from collections import Counter
import pandas as pd
from docx import Document
from openpyxl import load_workbook
from pathlib import Path
from copy import copy
from openpyxl.styles import Alignment
from openpyxl.worksheet.cell_range import CellRange
from decimal import InvalidOperation, ROUND_HALF_UP
import hashlib

COMPANY = '东莞市康铭光电科技有限公司'
IDENTITY = '4419360Q5N ' + COMPANY + '(91441900MA4UJ1P744)'

TEMPLATE_PATH = Path(__file__).resolve().parent / 'static' / 'kangming' / '出口报关单模板.xlsx'
VERSION = '20260930.4'

def _key(value):
    return re.sub(r'[\s:：()（）._/\\-]+', '', str(value)).upper()

def _required(value, label):
    if value is None or not str(value).strip():
        raise ValueError(label + '不能为空。')
    return str(value).strip()

def _number(value, label):
    try:
        number = Decimal(str(value).strip().replace(',', ''))
    except (InvalidOperation, ValueError):
        raise ValueError(label + '必须是有效数字。') from None
    if not number.is_finite() or number < 0:
        raise ValueError(label + '必须为非负有限数值。')
    return number

def _fixed(value):
    return format(Decimal(str(value)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP), '.2f')

def _fmt(value):
    return format(Decimal(str(value)), 'f').rstrip('0').rstrip('.') if '.' in str(value) else str(value)

def frame(content, label):
    sheets = pd.read_excel(BytesIO(content), sheet_name=None, header=None).values()
    matches = [s.fillna('') for s in sheets if any(label in str(v) for v in s.values.flat)]
    if len(matches) != 1:
        raise ValueError('未找到唯一的' + label + '工作表。')
    return matches[0]

def rule(name, model):
    key = _key(name)
    if key in {_key(n) for n in ['高尔夫球杆头毛坯配件/面板', '高尔夫球杆头毛坯配件/盖子', '高尔夫球杆头毛坯']}:
        return '9506390000', '享惠', '1.品牌类型：无品牌；2.出口享惠：{benefit}；3.型号：无型号；4.材质：不锈钢'
    if model in {'ZV25097-XM-3-S01', 'ZV26015-3-S01'}:
        return '8480719090', '享惠', '1.品牌类型：无品牌；2.出口享惠：{benefit}；3.用途：用于注塑模具上；4.适用材料：塑胶；5.品牌：无牌；6.型号：' + model + '；7.材质：1.2709热作；8.原理：注塑成型'
    if key in {'汽车控制单元外壳', '汽车运动控制单元外壳', '控制单元外壳'}:
        return '8708299000', '不享惠', '1.品牌类型：无品牌；2.出口享惠：{benefit}；3.用途：用于汽车运动中控制单元外壳；4.适用车型：通用；5.品牌：无牌；6.型号：无'
    raise ValueError('尚未配置商品【' + name + '】型号【' + model + '】的规则，请核对品名和型号。')

def read_data(invoice, packing):
    inv, pack = frame(invoice, 'INVOICE'), frame(packing, 'PACKING LIST')
    def text(df): return '\n'.join(str(v) for v in df.values.flat if str(v).strip())
    for df in (inv, pack):
        if COMPANY not in text(df): raise ValueError('发票、装箱单必须为康铭资料。')
    def contract(df):
        m = re.search(r'Contract\s*No\s*[:：]\s*([^\s]+)', text(df), re.I)
        if not m: raise ValueError('缺少合同号。')
        return m[1]
    contract_no = contract(inv)
    if contract_no != contract(pack): raise ValueError('发票和装箱单合同号不同。')
    def header(df):
        for i, row in df.iterrows():
            if '商品名称' in [str(v).strip() for v in row]:
                return i, {str(v).strip(): j for j, v in enumerate(row) if str(v).strip()}
        raise ValueError('缺少商品名称表头。')
    ih, ic = header(inv); ph, pc = header(pack)
    net_columns = [j for name,j in ic.items() if _key(name) in {'净重', '净重KG', 'NETWEIGHT', 'NETWEIGHTKG'}]
    if len(net_columns) != 1: raise ValueError('发票必须包含唯一的净重(KG)列。')
    items = []
    for _, row in inv.iloc[ih+1:].iterrows():
        name = str(row[ic['商品名称']]).strip()
        if not name: continue
        model = str(row[ic['规格']]).strip() or '无'
        if str(row[ic['品牌']]).strip() not in {'无', '无牌', '无品牌'}: raise ValueError('当前商品规则仅适用于无品牌。')
        code, benefit, elements = rule(name, model)
        qty = _number(row[ic['数量']], '数量'); amount = _number(row[ic['总价']], '总价')
        if qty <= 0: raise ValueError('数量必须大于零。')
        items.append(dict(name=name, model=model, code=code, qty=qty, amount=amount, price=amount/qty,
                          unit=_required(row[ic['单位']], '单位'), net_weight=_number(row[net_columns[0]], name+'净重'), benefit=benefit, elements=elements))
    if not items: raise ValueError('发票无商品。')
    totals = [row for _, row in inv.iterrows() if any('Total:' in str(v) for v in row)]
    if len(totals) != 1: raise ValueError('发票总计行不唯一。')
    total = _number(totals[0][ic['总价']], '发票合计')
    if abs(total - sum(i['amount'] for i in items)) > Decimal('.01'): raise ValueError('发票金额合计不一致。')
    meta = ' '.join(str(v) for v in totals[0])
    m = re.search(r'(C&F|CIF|FOB|CFR|EXW)\s+(.+?)\s+Total:', meta)
    currency = next((c for c in ['USD','EUR','CNY','HKD'] if c in meta), None)
    if not m or not currency: raise ValueError('发票总计行缺少成交方式、目的地或币种。')
    rows = []; summary = None
    for _, row in pack.iloc[ph+1:].iterrows():
        if any(str(v).strip() == '合计：' for v in row): summary=row; break
        if str(row[pc['商品名称']]).strip(): rows.append(row)
    if summary is None: raise ValueError('装箱单缺少合计行。')
    actual = Counter()
    for r in rows: actual[str(r[pc['商品名称']]).strip()] += _number(r[pc['数量']], '数量')
    expected = Counter()
    for i in items: expected[i['name']] += i['qty']
    if actual != expected: raise ValueError('发票与装箱单商品或数量不一致。')
    nw, gw = [_number(summary[pc[k]], k) for k in ['净重','毛重']]
    for key, total_weight in [('净重',nw),('毛重',gw)]:
        if abs(sum(_number(r[pc[key]],key) for r in rows)-total_weight)>Decimal('.01'): raise ValueError(key+'合计不一致。')
    packaging = re.search(r'包装件数：\s*(\d+)\s*([^\s]+)', text(pack))
    if not packaging: raise ValueError('装箱单缺少包装件数。')
    boxes = int(packaging[1])
    if boxes <= 0 or sum(_number(r[pc['箱数']], '箱数') for r in rows) != boxes: raise ValueError('箱数合计不一致。')
    if _number(summary[pc['数量']], '总数量') != sum(expected.values()): raise ValueError('总数量不一致。')
    if abs(sum(i['net_weight'] for i in items)-nw)>Decimal('.01'): raise ValueError('发票商品净重合计与装箱单净重不一致。')
    if nw<=0 or gw<nw: raise ValueError('净重或毛重不正确。')
    return dict(items=items, total_amount=total, currency=currency, incoterms=m[1], destination=m[2].strip(),
                contract_number=contract_no, net_weight=nw, gross_weight=gw, total_packages=boxes,
                pack_type=packaging[2], packages_text=packaging[0].split('：',1)[1].strip())

def create_kangming_export_declaration(data, inputs, template_path=None):
    """按康铭模板表头填写，保留净重列，动态移动页脚。"""
    wb = load_workbook(template_path or TEMPLATE_PATH)
    ws = wb.active
    required = {'项号', '商品编码', '商品名称及规格型号', '数量及单位', '净重KG',
                '单价', '总价', '币值', '原产国地区', '最终目的国地区', '境内货源地', '征免'}
    headers = None
    for row in ws:
        columns = {_key(c.value): c.column for c in row if c.value is not None}
        if '商品名称及规格型号' in columns:
            if not required.issubset(columns):
                raise ValueError('康铭模板表头缺失：' + '、'.join(sorted(required - columns.keys())))
            header_row, headers = row[0].row, columns
            break
    if headers is None: raise ValueError('康铭模板没有商品明细表头。')
    start = header_row + 1
    footer = next((c.row for row in ws.iter_rows(min_row=start) for c in row
                   if c.value and '特殊关系确认' in str(c.value)), None)
    if footer is None or footer <= start: raise ValueError('康铭模板缺少明细原型或页脚。')
    count = len(data['items'])
    if not count: raise ValueError('没有商品。')
    last_row, last_col = ws.max_row, max(headers.values())
    def snapshot(r):
        return [(ws.cell(r,c).value, copy(ws.cell(r,c)._style), copy(ws.cell(r,c).alignment))
                for c in range(1,last_col+1)]
    prototype = snapshot(start)
    tail = [(r, snapshot(r), ws.row_dimensions[r].height) for r in range(footer,last_row+1)]
    prototype_height = ws.row_dimensions[start].height or 54
    merges = [copy(m) for m in ws.merged_cells.ranges if m.min_row >= start]
    for m in merges: ws.unmerge_cells(str(m))
    ws.delete_rows(start, last_row-start+1)
    for r in list(ws.row_dimensions):
        if r >= start: del ws.row_dimensions[r]
    for n,item in enumerate(data['items']):
        r=start+n
        for c,(_,style,align) in enumerate(prototype,1):
            ws.cell(r,c)._style=copy(style); ws.cell(r,c).alignment=copy(align)
        values={'项号':n+1, '商品编码':item['code'], '商品名称及规格型号':item['name'],
                '数量及单位':_fmt(item['qty'])+item['unit'], '净重KG':float(_fixed(item['net_weight'])),
                '单价':float(_fixed(item['amount']/item['qty'])), '总价':float(_fixed(item['amount'])),
                '币值':{'USD':'美元','EUR':'欧元','CNY':'人民币','HKD':'港币'}[data['currency']],
                '原产国地区':'中国','最终目的国地区':data['destination'],'境内货源地':'东莞','征免':'照章征税'}
        for label,value in values.items():
            cell=ws.cell(r,headers[label],value)
            cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            if label in {'净重KG','单价','总价'}: cell.number_format='0.00'
        ws.row_dimensions[r].height=max(prototype_height,54)
    offset=start+count-footer
    for r,cells,height in tail:
        for c,(value,style,align) in enumerate(cells,1):
            cell=ws.cell(r+offset,c,value);cell._style=copy(style);cell.alignment=copy(align)
        ws.row_dimensions[r+offset].height=height
    for m in merges:
        if m.min_row>=footer:
            m.shift(row_shift=offset);ws.merge_cells(str(m))
    # 按现有标签查找汇总位置，避免新增净重列导致旧列号覆盖。
    updates = {
        '境内发货人':'境内发货人:'+IDENTITY,
        '境外收货人':'境外收货人：\n'+inputs['consignee'],
        '生产销售单位':'生产销售单位\n'+IDENTITY,
        '合同协议号':'合同协议号：'+data['contract_number'],
        '贸易国':'贸易国（地区）：\n'+inputs['trade_country'],
        '包装种类':'包装种类:'+inputs['pack_type'],
        '件数':f"件数:{data['total_packages']}件",
        '毛重':'毛重（千克):'+_fixed(data['gross_weight']),
        '净重':'净重（千克):'+_fixed(data['net_weight']),
        '成交方式':'成交方式:'+data['incoterms'],
        '运费':'运费:'+inputs.get('freight',''),
        '保费':'保费:'+inputs.get('insurance',''),
        '杂费':'杂费:'+inputs.get('other_fees',''),
        '标记唛码及备注':'标记唛码及备注 :'+data['packages_text'],
    }
    found=set()
    for row in ws.iter_rows(max_row=header_row-1):
        for cell in row:
            if cell.value is None: continue
            for label,value in updates.items():
                if str(cell.value).strip().startswith(label):
                    cell.value=value; found.add(label); break
    if not {'毛重','净重'}.issubset(found): raise ValueError('模板缺少毛重或净重汇总字段。')
    # 当前用户模板上方部分标签为空，恢复空白的固定海关字段。
    missing={'F3':'出口日期','H3':'申报日期','J3':'备案号','F4':'运输工具名称及航次号',
             'H4':'提运单号','F5':'征免性质\n一般征税','H5':'许可证号',
             'F6':'运抵国（地区）\n'+data['destination'],'H6':'指运港','J6':'离境口岸'}
    for address,value in missing.items():
        if ws[address].value is None: ws[address]=value
    if '杂费' not in found: ws['L7']='杂费:'+inputs.get('other_fees','')
    for row in ws.iter_rows(max_row=header_row-1):
        for cell in row:
            if cell.value is not None: cell.alignment=Alignment(wrap_text=True,vertical='center')
    from openpyxl.utils import get_column_letter
    ws.print_area=f'A1:{get_column_letter(last_col)}{last_row+offset}'
    ws.print_title_rows=f'1:{header_row}'
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=1 if count<=5 else 0
    out=BytesIO();wb.save(out);wb.close()
    return out.getvalue()


def generate(data, inputs, invoice, packing):
    for key in ['consignee','trade_country','pack_type']:
        _required(inputs.get(key), key)
    benefit=inputs.get('benefit')
    if benefit not in {'享惠','不享惠'}: raise ValueError('请选择本批次出口享惠状态。')
    doc = Document(); doc.add_heading('商品申报要素清单', 0); doc.add_paragraph(COMPANY)
    for n, i in enumerate(data['items'],1):
        doc.add_paragraph(f"{n}、商品编码：{i['code']}  {i['name']}")
        doc.add_paragraph(i['elements'].format(benefit=benefit))
    out=BytesIO(); doc.save(out)
    customs=create_kangming_export_declaration(data, inputs)
    stamp=datetime.now(ZoneInfo('Asia/Shanghai')).strftime('%Y%m%d')
    docs={f'康铭_申报要素_{stamp}.docx':out.getvalue(), f'康铭_出口报关单_{stamp}.xlsx':customs}
    for label, content in [('发票',invoice),('装箱单',packing)]:
        ext='xlsx' if content.startswith(b'PK\x03\x04') else 'xls'
        docs[f'康铭_{label}_{stamp}.{ext}']=content
    z=BytesIO()
    with zipfile.ZipFile(z,'w',zipfile.ZIP_DEFLATED) as archive:
        for name,content in docs.items(): archive.writestr(name,content)
    return z.getvalue()

def render():
    import streamlit as st
    st.header('东莞康铭 · 出口单证')
    invoice=st.file_uploader('上传发票', type=['xls','xlsx'],key='km_invoice')
    packing=st.file_uploader('上传装箱单', type=['xls','xlsx'],key='km_packing')
    if not invoice or not packing:
        st.session_state.pop('km_archive',None)
        return
    try: data=read_data(invoice.getvalue(),packing.getvalue())
    except Exception as exc:
        st.session_state.pop('km_archive',None)
        st.error(str(exc)); return
    st.success(f"读取成功：{len(data['items'])} 项商品，{data['total_amount']:.2f} {data['currency']}，{data['total_packages']} 箱。")
    benefit=st.selectbox('出口享惠（本批所有商品）',['享惠','不享惠'],key='km_benefit_global')
    inputs={'benefit':benefit}
    for key,label,default in [('consignee','境外收货人','CANGMING 3D TECH CO.,LIMITED'),('trade_country','贸易国','中国香港'),('pack_type','包装种类',data['pack_type']),('freight','运费',''),('insurance','保费',''),('other_fees','杂费','')]:
        inputs[key]=st.text_input(label,value=default,key='km_'+key).strip()
    st.caption('康铭版本 '+VERSION+' · 金额及商品净重取自发票，享惠状态统一应用。')
    stamp=datetime.now(ZoneInfo('Asia/Shanghai')).strftime('%Y%m%d')
    signature=(hashlib.sha256(invoice.getvalue()).hexdigest(),hashlib.sha256(packing.getvalue()).hexdigest(),tuple(sorted(inputs.items())),stamp)
    if st.session_state.get('km_signature')!=signature:
        st.session_state.pop('km_archive',None)
    if st.button('生成康铭单证',key='km_generate'):
        st.session_state.pop('km_archive',None)
        try:
            st.session_state['km_archive']=generate(data,inputs,invoice.getvalue(),packing.getvalue())
            st.session_state['km_signature']=signature
        except Exception as exc: st.error(str(exc))
    if st.session_state.get('km_archive') is not None:
        st.download_button('下载康铭单证 ZIP',st.session_state['km_archive'],f'康铭_单证_{stamp}.zip','application/zip',key='km_download')
