#!/usr/bin/env python3
"""Synthetic tests only: these amounts are not current commercial quotations."""
import copy
import json
import tempfile
import subprocess
import sys
import unittest
from pathlib import Path
from openpyxl import load_workbook
from build_quote import build_workbook, template_data, validate, HEADERS


def row(rid, group='产品', mode='unit'):
    return {'id':rid,'group':group,'product_line':'测试产品','service':'模拟服务'+rid,'selection':'必选',
            'functions':'合成测试内容','status':'confirmed','mode':mode,'unit_price':'100','quantity':'2','unit':'个',
            'discount':None if mode=='net' else '0.8','periods':'1','source_id':'source1','source_ref':'合成测试定位',
            'sku_key':'test:'+rid,'price_current':True,'quote_authorized':True,'discount_authorized':True,'scope_confirmed':True,
            'evidence':{k:'合成测试证据' for k in ['scope','price','quantity','discount','periods','rate','basis','zero']}}


def fixture():
    d=template_data();d['title']='合成测试报价，不用于客户';d['rows']=[row('p'),row('i','实施'),row('m','运维')]
    d['source_checks']=[{'source_id':'source'+str(i),'state':'matched'} for i in range(1,5)]
    d['commercial']={'tax_basis':'测试含税口径','quote_date':'2026-11-10','valid_until':'2026-11-30','conditions_confirmed':True}
    d['confirmed_business']=True
    return d


class QuoteTests(unittest.TestCase):
    def test_empty_template_has_all_dimensions(self):
        wb,r=build_workbook(template_data());self.assertEqual([wb['报价明细'].cell(2,i).value for i in range(1,10)],HEADERS)
        self.assertEqual(set(r['subtotals']),{'产品','实施','运维'});self.assertIsNone(r['total']);self.assertFalse(r['formal_ready'])
        for rr in [3,5,7]:self.assertIsNone(wb['报价明细'].cell(rr,5).value)
    def test_normal_price(self):
        rows,r=validate(fixture());self.assertEqual(r['amounts']['p'],'160.0');self.assertEqual(r['total'],'480.0')
    def test_missing_group_remains_blank(self):
        d=fixture();d['confirmed_business']=False;d['rows']=d['rows'][:1];rows,r=validate(d)
        self.assertEqual(len(rows),3);self.assertIsNone(r['subtotals']['实施']);self.assertIsNone(r['total'])
    def test_unknown_with_numeric_values_cannot_release_total(self):
        d=fixture();d['confirmed_business']=False;d['rows'][0]['status']='unknown';_,r=validate(d)
        self.assertIsNone(r['amounts']['p']);self.assertIsNone(r['subtotals']['产品']);self.assertIsNone(r['total'])
    def test_empty_discount_not_one(self):
        d=fixture();d['rows'][0]['discount']=None
        with self.assertRaises(ValueError):validate(d)
    def test_net_price_not_discounted_twice(self):
        d=fixture();d['rows'][0]=row('p',mode='net');_,r=validate(d);self.assertEqual(r['amounts']['p'],'200')
        d['rows'][0]['discount']='0.8'
        with self.assertRaises(ValueError):validate(d)
    def test_rate_uses_only_prior_basis(self):
        d=fixture();m=d['rows'][2];m.update(mode='percentage',rate='0.2',basis_rows=['p'],unit='年');_,r=validate(d);self.assertEqual(r['amounts']['m'],'32.00')
        m['basis_rows']=['m']
        with self.assertRaises(ValueError):validate(d)
    def test_percentage_missing_rate(self):
        d=fixture();d['rows'][2].update(mode='percentage',basis_rows=['p'],rate=None)
        with self.assertRaises(ValueError):validate(d)
    def test_evidence_free_zero(self):
        d=fixture();d['rows'][2].update(status='included',unit_price=None,quantity=None,discount=None,periods=None)
        _,r=validate(d);self.assertEqual(r['subtotals']['运维'],'0');self.assertTrue(r['formal_ready'])
        d['rows'][2]['evidence'].pop('zero')
        with self.assertRaises(ValueError):validate(d)
    def test_optional_not_selected_has_empty_amount(self):
        d=fixture();d['rows'][2].update(status='not_selected');_,r=validate(d)
        self.assertIsNone(r['amounts']['m']);self.assertEqual(r['subtotals']['运维'],'0')
    def test_unsupported_status(self):
        d=fixture();d['rows'][0]['status']='confimed'
        with self.assertRaises(ValueError):validate(d)
    def test_missing_source_prevents_formal(self):
        d=fixture();d['source_checks']=d['source_checks'][:3]
        with self.assertRaises(ValueError):validate(d)
    def test_blocked_source_cannot_formal(self):
        d=fixture();d['source_checks'][1]['state']='blocked'
        with self.assertRaises(ValueError):validate(d)
    def test_source_and_price_authority(self):
        for key in ['price_current','quote_authorized','discount_authorized','scope_confirmed']:
            d=fixture();d['rows'][0][key]=False
            with self.assertRaises(ValueError):validate(d)
    def test_duplicate_sku(self):
        d=fixture();d['rows'][1]['sku_key']=d['rows'][0]['sku_key']
        with self.assertRaises(ValueError):validate(d)
    def test_missing_commercial_total_blank(self):
        d=fixture();d['confirmed_business']=False;d['commercial']['tax_basis']=None;_,r=validate(d);self.assertIsNone(r['total'])
        d['confirmed_business']=True
        with self.assertRaises(ValueError):validate(d)
    def test_invalid_numbers(self):
        for v in ['NaN','Infinity','-1',True,'1e30']:
            d=fixture();d['rows'][0]['unit_price']=v
            with self.assertRaises(ValueError):validate(d)
        d=fixture();d['rows'][0]['discount']='6'
        with self.assertRaises(ValueError):validate(d)
    def test_formulas_are_written_and_text_safe(self):
        d=fixture();d['rows'][0]['functions']='=HYPERLINK("https://invalid.test")';wb,r=build_workbook(d)
        self.assertEqual(wb['报价明细']['D3'].data_type,'s');self.assertEqual(wb['报价明细']['I3'].data_type,'f')
        self.assertIn('ISNUMBER',wb['报价明细']['I3'].value);self.assertIn('报价条件',wb['报价明细']['I3'].value)
        self.assertEqual(wb['报价明细']['A1'].fill.fgColor.rgb[-6:],'477CFB');self.assertEqual(wb['报价明细']['A4'].fill.fgColor.rgb[-6:],'FFF24D')
    def test_many_rows_and_totals(self):
        d=fixture();d['rows']=[row(str(i)) for i in range(15)]+[row('i','实施'),row('m','运维')];wb,r=build_workbook(d)
        self.assertEqual(len(r['row_cells']),17);self.assertEqual(r['total'],'2720.0')
        self.assertIn('I18',wb['报价明细'][r['total_cell'].split('!')[1]].value)
    def test_quote_dates_and_selection(self):
        d=fixture();d['commercial']['valid_until']='2026-11-01'
        with self.assertRaises(ValueError):validate(d)
        d=fixture();d['rows'][0]['selection']='默认购买'
        with self.assertRaises(ValueError):validate(d)
    def test_missing_period_not_one(self):
        d=fixture();d['rows'][0]['periods']=None
        with self.assertRaises(ValueError):validate(d)
    def test_no_input_mutation(self):
        d=fixture();original=copy.deepcopy(d);validate(d);self.assertEqual(d,original)
    def test_save_reload(self):
        with tempfile.TemporaryDirectory() as td:
            wb,r=build_workbook(fixture());path=Path(td)/'quote.xlsx';wb.save(path);saved=load_workbook(path)
            self.assertEqual(saved['报价明细']['I3'].data_type,'f');self.assertEqual(len(saved.sheetnames),2)
    def test_cli_refuses_overwrite(self):
        with tempfile.TemporaryDirectory() as td:
            output=Path(td)/'template.xlsx';cmd=[sys.executable,'-B',str(Path(__file__).with_name('build_quote.py')),'--template','--output',str(output)]
            first=subprocess.run(cmd,capture_output=True);self.assertEqual(first.returncode,0,first.stderr)
            second=subprocess.run(cmd,capture_output=True);self.assertEqual(second.returncode,2)

if __name__=='__main__':unittest.main(verbosity=2)
