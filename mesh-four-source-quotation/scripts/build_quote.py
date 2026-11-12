#!/usr/bin/env python3
"""Four-source quotation: validated inputs, blank-safe formulas and screenshot-like layout."""
import argparse
import json
import math
import sys
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VERSION = '1.0.0'
GROUPS = ('产品', '实施', '运维')
SOURCES = {'source1', 'source2', 'source3', 'source4'}
STATES = {'confirmed', 'free', 'included', 'not_selected', 'not_applicable', 'unknown'}
HEADERS = ['产品线', '服务项', '必选/可选', '功能清单', '单价（元）', '数量', '单位', '折扣', '折后价（元）']
BLUE, YELLOW, BLACK = '477CFB', 'FFF24D', '000000'


def dec(value, name):
    if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
        raise ValueError(name + ': invalid numeric type')
    try:
        x = Decimal(str(value))
    except InvalidOperation as exc:
        raise ValueError(name + ': invalid number') from exc
    if not x.is_finite() or x < 0 or (x and abs(x.adjusted()) > 12):
        raise ValueError(name + ': expected finite nonnegative supported number')
    return x


def text(cell, value):
    """Force untrusted source/customer content to remain literal, not a formula."""
    if value is None:
        return
    cell.value = str(value)
    cell.data_type = 's'


def template_data():
    return {'title': '按功能模块费用明细—空白报价模板', 'currency': 'CNY',
            'confirmed_business': False, 'commercial': {},
            'source_checks': [{'source_id': x, 'state': 'blocked'} for x in sorted(SOURCES)],
            'rows': [{'id': 'blank_' + str(i), 'group': group, 'status': 'unknown'}
                     for i, group in enumerate(GROUPS)]}


def validate(data):
    if not isinstance(data, dict) or data.get('currency') != 'CNY':
        raise ValueError('Input must be an object with currency=CNY')
    checks = data.get('source_checks', [])
    if not isinstance(checks, list) or any(not isinstance(x, dict) for x in checks):
        raise ValueError('Invalid source_checks')
    if len(checks) != 4 or {x.get('source_id') for x in checks} != SOURCES:
        raise ValueError('All four sources must be checked exactly once')
    if any(x.get('state') not in {'matched', 'no_match', 'blocked', 'conflict'} for x in checks):
        raise ValueError('Unknown source check state')
    source_complete = all(x['state'] in {'matched', 'no_match'} for x in checks)
    rows = data.get('rows')
    if not isinstance(rows, list) or any(not isinstance(x, dict) for x in rows):
        raise ValueError('rows must be an array of objects')
    rows = [dict(x) for group in GROUPS for x in rows if x.get('group') == group]
    if len(rows) != len(data['rows']):
        raise ValueError('Unknown group; only 产品/实施/运维 allowed')
    for group in GROUPS:
        if not any(x['group'] == group for x in rows):
            rows.append({'id': 'missing_' + group, 'group': group, 'status': 'unknown'})
    rows = [x for group in GROUPS for x in rows if x['group'] == group]
    ids = [x.get('id') for x in rows]
    if any(not isinstance(x, str) or not x for x in ids) or len(set(ids)) != len(ids):
        raise ValueError('Nonempty unique row IDs required')
    by_id, amounts, sku_seen, gaps = {}, {}, set(), []
    with localcontext() as context:
        context.prec = 40
        for row in rows:
            rid, state = row['id'], row.get('status')
            if state not in STATES:
                raise ValueError(rid + ': invalid state')
            mode = row.get('mode', 'unit')
            if row.get('selection') not in {None, '必选', '可选'}:
                raise ValueError(rid + ': selection must be 必选/可选 or null')
            if mode not in {'unit', 'net', 'percentage'}:
                raise ValueError(rid + ': unsupported price mode; retain unknown pending specialized formula')
            row['mode'] = mode
            evidence = row.get('evidence', {})
            if not isinstance(evidence, dict):
                raise ValueError(rid + ': evidence must be object')
            amount = None
            numeric = {}
            for key in ['unit_price', 'quantity', 'discount', 'periods', 'rate']:
                if row.get(key) is not None:
                    numeric[key] = dec(row[key], rid + '.' + key)
                    row[key] = str(numeric[key])
            if 'discount' in numeric and numeric['discount'] > 1:
                raise ValueError(rid + ': discount must be a factor from 0 to 1')
            if 'rate' in numeric and numeric['rate'] > 1:
                raise ValueError(rid + ': rate must be from 0 to 1')
            if state in {'confirmed', 'free', 'included'}:
                if row.get('source_id') not in SOURCES or not row.get('source_ref') or not row.get('sku_key'):
                    raise ValueError(rid + ': missing valid source/row identity')
                if row.get('scope_confirmed') is not True or row.get('quote_authorized') is not True:
                    raise ValueError(rid + ': missing scope confirmation/quotation authority')
                if not evidence.get('scope'):
                    raise ValueError(rid + ': scope evidence missing')
                sku = row['sku_key']
                if sku in sku_seen:
                    raise ValueError(rid + ': duplicate selected SKU; resolve package overlap')
                sku_seen.add(sku)
                if not row.get('service') or not row.get('product_line'):
                    raise ValueError(rid + ': product and service names required')
                if state in {'free', 'included'}:
                    if not evidence.get('zero'):
                        raise ValueError(rid + ': zero amount needs applicable free/included evidence')
                    amount = Decimal(0)
                else:
                    if row.get('price_current') is not True or row.get('discount_authorized') is not True:
                        raise ValueError(rid + ': current applicable price and factor authority required')
                    if not row.get('unit') or not evidence.get('price') or not evidence.get('quantity') or not evidence.get('periods'):
                        raise ValueError(rid + ': price, quantity/unit and period evidence missing')
                    periods = numeric.get('periods')
                    if periods is None or periods <= 0:
                        raise ValueError(rid + ': explicitly confirmed positive periods required')
                    if mode == 'percentage':
                        basis = row.get('basis_rows')
                        if not isinstance(basis, list) or not basis or len(set(basis)) != len(basis):
                            raise ValueError(rid + ': nonempty unique basis_rows required')
                        if any(x not in amounts or amounts[x] is None for x in basis):
                            raise ValueError(rid + ': base must be prior confirmed rows, not self/future/unknown')
                        if 'rate' not in numeric or not evidence.get('basis') or not evidence.get('rate'):
                            raise ValueError(rid + ': rate and base evidence missing')
                        amount = sum(amounts[x] for x in basis) * numeric['rate'] * periods
                    else:
                        if any(k not in numeric for k in ['unit_price', 'quantity']) or numeric['quantity'] <= 0:
                            raise ValueError(rid + ': positive confirmed quantity and price required')
                        if mode == 'net':
                            if row.get('discount') is not None:
                                raise ValueError(rid + ': net price cannot be discounted again')
                            factor = Decimal(1)
                        else:
                            if numeric.get('discount') is None or numeric['discount'] <= 0 or not evidence.get('discount'):
                                raise ValueError(rid + ': confirmed nonzero factor evidence required')
                            factor = numeric['discount']
                        amount = numeric['unit_price'] * numeric['quantity'] * factor * periods
            elif state in {'not_selected', 'not_applicable'}:
                if not evidence.get('scope') or row.get('scope_confirmed') is not True:
                    raise ValueError(rid + ': explicit exclusion evidence required')
            else:
                gaps.append({'row_id': rid, 'group': row['group'], 'missing': [k for k in ['unit_price', 'quantity', 'unit', 'discount', 'periods'] if row.get(k) is None], 'reason': row.get('gap', '资料或适用条件未确认')})
            row['_ready'] = state != 'unknown'
            row['_amount'] = str(amount) if amount is not None else None
            amounts[rid] = amount
            by_id[rid] = row
        subtotals = {}
        for group in GROUPS:
            members = [x for x in rows if x['group'] == group]
            subtotals[group] = None if any(x['status'] == 'unknown' for x in members) else sum(amounts[x['id']] or Decimal(0) for x in members)
        commercial = data.get('commercial', {})
        if not isinstance(commercial, dict):
            raise ValueError('commercial must be object')
        gate = data.get('confirmed_business') is True and commercial.get('conditions_confirmed') is True and source_complete
        gate = gate and all(isinstance(commercial.get(k), str) and commercial[k].strip() for k in ['tax_basis', 'quote_date', 'valid_until'])
        if gate:
            try:
                start = date.fromisoformat(commercial['quote_date'])
                end = date.fromisoformat(commercial['valid_until'])
            except ValueError as exc:
                raise ValueError('Quote dates must be ISO dates') from exc
            if end < start:
                raise ValueError('Quote validity cannot end before its issue date')
        if data.get('confirmed_business') is True and (not gate or any(x is None for x in subtotals.values())):
            raise ValueError('Formal business flag requires complete sources, commercial fields and three dimensions')
        total = sum(subtotals.values()) if gate and all(x is not None for x in subtotals.values()) else None
    return rows, {'version': VERSION, 'formal_ready': bool(total is not None), 'source_complete': source_complete,
                  'subtotals': {k: str(v) if v is not None else None for k, v in subtotals.items()},
                  'total': str(total) if total is not None else None, 'gaps': gaps,
                  'amounts': {k: str(v) if v is not None else None for k, v in amounts.items()}}


def build_workbook(data):
    rows, report = validate(data)
    wb = Workbook()
    ws = wb.active; ws.title = '报价明细'
    controls = wb.create_sheet('报价条件')
    for sheet in wb:
        sheet.sheet_view.showGridLines = False
    ws.merge_cells('A1:I1');text(ws['A1'], data.get('title') or '按功能模块费用明细')
    for i, h in enumerate(HEADERS, 1):text(ws.cell(2, i), h)
    control_values = [(1,'报价条件与确认状态',''),(2,'商务条件已确认',report['formal_ready']),
                      (3,'四源覆盖已完成',report['source_complete']),(4,'币种','CNY'),
                      (5,'报价日期',data.get('commercial',{}).get('quote_date')),
                      (6,'含税口径',data.get('commercial',{}).get('tax_basis')),
                      (7,'有效期截止',data.get('commercial',{}).get('valid_until')),
                      (8,'说明','修改价格、折扣或范围后需重新核验并将状态改为unknown，确认后恢复。')]
    for r, label, value in control_values:
        text(controls.cell(r,1),label)
        if isinstance(value,bool):controls.cell(r,2,value)
        else:text(controls.cell(r,2),value)
    for i,h in enumerate(['行项','状态','计费周期数','运维费率','行确认','单位口径'],1):text(controls.cell(10,i),h)
    row_map, control_map, subtotal_rows = {}, {}, {}
    curr, cidx = 3, 11
    for group in GROUPS:
        group_start=curr
        for row in [x for x in rows if x['group']==group]:
            rid=row['id'];row_map[rid]=curr;control_map[rid]=cidx
            values=[row.get('product_line') or (group if group!='产品' else None),row.get('service'),row.get('selection'),row.get('functions'),row.get('unit_price'),row.get('quantity'),row.get('unit'),row.get('discount')]
            for col,value in enumerate(values,1):
                if col in {5,6,8} and value is not None:
                    num=float(Decimal(str(value)))
                    if not math.isfinite(num):raise ValueError('Number beyond Excel range')
                    ws.cell(curr,col,num)
                else:text(ws.cell(curr,col),value)
            if row['mode']=='net':text(ws.cell(curr,4),(row.get('functions') or '')+'\n单价为已折后单价。')
            if row['mode']=='percentage':text(ws.cell(curr,4),(row.get('functions') or '')+'\n按已确认基数及费率计费。')
            text(controls.cell(cidx,1), row.get('service') or group+'待确认行')
            text(controls.cell(cidx,2),row['status']);controls.cell(cidx,5,row['_ready'])
            for col,key in [(3,'periods'),(4,'rate')]:
                if row.get(key) is not None:controls.cell(cidx,col,float(Decimal(row[key])))
            text(controls.cell(cidx,6),row.get('unit'))
            state=f"'报价条件'!B{cidx}";period=f"'报价条件'!C{cidx}";ready=f"'报价条件'!E{cidx}"
            if row['mode']=='percentage':
                basis=row.get('basis_rows',[]);refs=[f'I{row_map[x]}' for x in basis]
                numeric=f'COUNT({",".join(refs)})={len(refs)}' if refs else 'FALSE'
                expr='('+('+'.join(refs) or '0')+f")*'报价条件'!D{cidx}*{period}"
                valid=f'AND({numeric},ISNUMBER(\'报价条件\'!D{cidx}),\'报价条件\'!D{cidx}>=0,\'报价条件\'!D{cidx}<=1,ISNUMBER({period}),{period}>0)'
            else:
                extra=f',ISNUMBER(H{curr}),H{curr}>0,H{curr}<=1' if row['mode']=='unit' else f',H{curr}=""'
                valid=f'AND(ISNUMBER(E{curr}),E{curr}>=0,ISNUMBER(F{curr}),F{curr}>0,ISNUMBER({period}),{period}>0{extra})'
                expr=f'E{curr}*F{curr}'+(f'*H{curr}' if row['mode']=='unit' else '')+f'*{period}'
            formula=f'=IF({ready}<>TRUE,"",IF(OR({state}="free",{state}="included"),0,IF(AND({state}="confirmed",{valid}),{expr},"")))'
            ws.cell(curr,9,formula)
            curr+=1;cidx+=1
        group_end=curr-1;sumrow=curr;subtotal_rows[group]=curr
        ws.merge_cells(start_row=curr,start_column=1,end_row=curr,end_column=8)
        text(ws.cell(curr,1),group+('费用' if group=='产品' else '服务')+'合计（折后价）')
        control_start=cidx-(group_end-group_start+1);rg=f"'报价条件'!B{control_start}:B{cidx-1}";vals=f'I{group_start}:I{group_end}'
        recognized=f'COUNTIF({rg},"confirmed")+COUNTIF({rg},"free")+COUNTIF({rg},"included")+COUNTIF({rg},"not_selected")+COUNTIF({rg},"not_applicable")'
        expected=f'COUNTIF({rg},"confirmed")+COUNTIF({rg},"free")+COUNTIF({rg},"included")'
        ws.cell(curr,9,f'=IF(AND({recognized}=ROWS({rg}),COUNT({vals})={expected}),SUM({vals}),"")')
        for col in range(1,10):ws.cell(curr,col).fill=PatternFill('solid',fgColor=YELLOW);ws.cell(curr,col).font=Font(name='微软雅黑',size=11,bold=True)
        # Contiguous equal product labels only; preserve distinct names.
        start=group_start
        for rr in range(group_start+1,group_end+2):
            if rr==group_end+1 or ws.cell(rr,1).value!=ws.cell(start,1).value:
                if rr-start>1 and ws.cell(start,1).value:ws.merge_cells(start_row=start,end_row=rr-1,start_column=1,end_column=1)
                start=rr
        curr+=1
    ws.merge_cells(start_row=curr,start_column=1,end_row=curr,end_column=8);text(ws.cell(curr,1),'总价')
    refs=','.join('I'+str(subtotal_rows[g]) for g in GROUPS)
    ws.cell(curr,9,f'=IF(AND(\'报价条件\'!B2=TRUE,\'报价条件\'!B3=TRUE,COUNT({refs})=3),SUM({refs}),"")')
    total_row=curr
    for col in range(1,10):ws.cell(curr,col).fill=PatternFill('solid',fgColor=YELLOW)
    ws.merge_cells(start_row=curr+1,start_column=1,end_row=curr+1,end_column=9)
    ws.cell(curr+1,1,f'=IF(ISNUMBER(I{curr}),"正式报价：条件和三维金额已确认","待确认草稿：未确认字段留空，总价未定")')
    ws.merge_cells(start_row=curr+2,start_column=1,end_row=curr+2,end_column=9)
    gaps='；'.join(group+'内容/金额待确认' for group in GROUPS if report['subtotals'][group] is None)
    text(ws.cell(curr+2,1),gaps or ('商务条件待确认' if not report['formal_ready'] else '按已确认币种、税口径及有效期执行。'))
    line=Side(style='thin',color=BLACK)
    for rowcells in ws.iter_rows(min_row=1,max_row=curr,max_col=9):
        for cell in rowcells:
            cell.border=Border(left=line,right=line,top=line,bottom=line)
            cell.alignment=Alignment(horizontal='left' if cell.column in {1,2,4} else 'center',vertical='center',wrap_text=True)
            if cell.row<=2:cell.fill=PatternFill('solid',fgColor=BLUE);cell.font=Font(name='微软雅黑',bold=True,color='FFFFFF',size=12)
            elif cell.row in list(subtotal_rows.values())+[total_row]:cell.font=Font(name='微软雅黑',bold=True,color=BLACK,size=11)
            else:cell.font=Font(name='微软雅黑',color='008000' if cell.data_type=='f' else '0000FF' if cell.column in {5,6,8} else BLACK,size=11)
            if cell.column in {5,9}:cell.number_format='¥#,##0.00;[Red](¥#,##0.00);¥0.00'
            if cell.column==6:cell.number_format='#,##0.##'
            if cell.column==8:cell.number_format='0.00'
    for col,width in zip('ABCDEFGHI',[23,31,15,57,17,14,18,13,20]):ws.column_dimensions[col].width=width
    for rownum in range(1,curr+3):
        ws.row_dimensions[rownum].height=26 if rownum in [1,2,total_row] or rownum in subtotal_rows.values() else max(105, min(400, 22 * max(len(str(ws.cell(rownum,4).value or '').splitlines()), math.ceil(len(str(ws.cell(rownum,4).value or ''))/32)))) if rownum in row_map.values() else 32
    for rowcells in controls:
        for cell in rowcells:
            cell.font=Font(name='微软雅黑',size=11,color='0000FF' if cell.column>1 else BLACK)
            cell.alignment=Alignment(vertical='center',wrap_text=True)
    controls.column_dimensions['A'].width=32;controls.column_dimensions['B'].width=55
    for col in 'CDEF':controls.column_dimensions[col].width=18
    for rr in [1,10]:
        for cell in controls[rr]:cell.fill=PatternFill('solid',fgColor=BLUE);cell.font=Font(name='微软雅黑',bold=True,color='FFFFFF',size=11)
    controls.row_dimensions[8].height=45
    report.update(row_cells={rid:'报价明细!I'+str(rr) for rid,rr in row_map.items()},subtotal_cells={g:'报价明细!I'+str(rr) for g,rr in subtotal_rows.items()},total_cell='报价明细!I'+str(total_row))
    return wb,report


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    group=parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--input',type=Path);group.add_argument('--template',action='store_true')
    parser.add_argument('--output',type=Path,required=True);parser.add_argument('--report',type=Path)
    args=parser.parse_args()
    try:
        if args.output.exists() or (args.report and args.report.exists()):raise ValueError('Refusing to overwrite existing output/report')
        data=template_data() if args.template else json.loads(args.input.read_text(encoding='utf-8-sig'))
        wb,report=build_workbook(data);wb.save(args.output)
        if args.report:args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps({'ok':True,'file':str(args.output),'formal_ready':report['formal_ready'],'total':report['total']},ensure_ascii=False))
        return 0
    except (ValueError,TypeError,KeyError,OSError,InvalidOperation) as exc:
        print('Quotation build failed: '+str(exc),file=sys.stderr);return 2

if __name__=='__main__':sys.exit(main())
