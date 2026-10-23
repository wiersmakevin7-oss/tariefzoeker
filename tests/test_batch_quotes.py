import ast
import sys
import types
import unittest
from pathlib import Path
from io import BytesIO

import openpyxl
from batch_quotes import enrich_workbook, numeric, add_pallet_scale


def load_engine():
    """Load existing pricing definitions without running Streamlit UI or secrets."""
    path = Path(__file__).resolve().parents[1]/'app.py'
    tree = ast.parse(path.read_text(encoding='utf-8-sig'))
    keep = []
    for node in tree.body:
        if isinstance(node,ast.Expr) and isinstance(node.value,ast.Call) and isinstance(node.value.func,ast.Attribute) and node.value.func.attr=='set_page_config':
            break
        keep.append(node)
    module = types.ModuleType('batch_test_engine')
    module.__file__ = str(path)
    sys.modules[module.__name__] = module
    exec(compile(ast.Module(body=keep,type_ignores=[]),str(path),'exec'),module.__dict__)
    return module


class BatchTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.engine = load_engine()

    def workbook(self):
        book=openpyxl.Workbook()
        ws=book.active
        ws.title='Shipments'
        ws.append(['Postcode losadres','Totaalgewicht','Totaal laadmeter','Aantal Pallets','Bestaand'])
        ws.append(['01234',300,0.4,1,'=1+1'])
        ws.append(['01234',900,0.4,1,'behouden'])
        ws.append(['01234',None,0.4,1,'ongeldig'])
        ws.append(['99999',300,0.4,1,'geen match'])
        ws.append(['01234',300,0.4,1,'laatste'])
        book.create_sheet('Overig')['A1']='Behouden'
        buffer=BytesIO()
        book.save(buffer)
        return buffer.getvalue()

    def rates(self):
        R=self.engine.RateRow
        return [R('FR','01',500,0.4,None,None,100,None,'XPO','test.xlsx'),
                R('FR','01',1000,0.8,None,None,150,None,'XPO','test.xlsx'),
                R('FR','02',500,0.4,None,None,100,None,'XPO','test.xlsx'),
                R('FR','01',1000,1,None,None,120,None,'Van Heugten','test.xlsx'),
                R('FR','02',1000,1,None,None,120,None,'Van Heugten','test.xlsx'),
                R('AT','6',1000,1,None,None,50,None,'Drost','at.xlsx')]

    def run_batch(self, include=True, kind='Onbekend', rates=None):
        return enrich_workbook(self.workbook(),'Shipments',{'postcode':1,'kg':2,'ldm':3,'pallets':4},'FR',kind,
                               self.rates() if rates is None else rates,{'XPO':26.4},{'XPO':5},10,include,self.engine.find_rates)

    def test_all_carriers_and_preservation(self):
        raw,preview,counts=self.run_batch()
        book=openpyxl.load_workbook(BytesIO(raw))
        ws=book['Shipments']
        self.assertEqual(ws['E2'].value,'=1+1')
        self.assertEqual(book['Overig']['A1'].value,'Behouden')
        self.assertEqual(ws['A2'].value,'01234')
        self.assertEqual(counts,{'regels':5,'met_tarief':3,'zonder_tarief':1,'ongeldig':1})
        row=preview[0]
        self.assertEqual(ws.max_column,9)
        self.assertEqual(row['XPO - Totale inkoop EUR'],131.4)
        self.assertEqual(row['XPO - Verkoop EUR'],144.54)
        self.assertEqual(row['Van Heugten - Totale inkoop EUR'],120)
        self.assertFalse(any('Drost' in key for key in row))
        self.assertEqual(preview[1]['XPO - Totale inkoop EUR'],197.1)
        self.assertTrue(preview[2]['XPO - Totale inkoop EUR'].startswith('Ongeldige invoer'))
        self.assertEqual(preview[3]['XPO - Totale inkoop EUR'],'Geen tarief')
        self.assertEqual(preview[4]['XPO - Totale inkoop EUR'],131.4)

    def test_margin_optional(self):
        _,preview,_=self.run_batch(False)
        self.assertNotIn('XPO - Verkoop EUR',preview[0])
        self.assertEqual(preview[0]['XPO - Totale inkoop EUR'],131.4)

    def test_unknown_pallets_skip_count_brackets(self):
        R=self.engine.RateRow
        rates=[R('FR','FR',1000,1,1,None,100,None,'XPO','test.xlsx')]
        _,preview,_=self.run_batch(rates=rates)
        self.assertEqual(preview[0]['XPO - Totale inkoop EUR'],'Geen tarief')
        _,preview,_=self.run_batch(kind='Europallet',rates=rates)
        self.assertEqual(preview[0]['XPO - Totale inkoop EUR'],131.4)

    def test_numbers(self):
        self.assertEqual(numeric('1.400,50'),1400.5)
        self.assertEqual(numeric('0,4'),0.4)
        for bad in [None,True,'NaN','inf',0,-1,'tekst']:
            with self.assertRaises(ValueError): numeric(bad)

    def test_scale_destinations_and_quantities(self):
        R=self.engine.RateRow
        rates=[R('FR','FR',qty*300,round(qty*0.4,2),qty,None,qty*100,None,'XPO','source') for qty in range(1,5)]
        rates.append(R('AT','6',1000,1,None,None,50,None,'Drost','source'))
        output,counts=add_pallet_scale(self.workbook(),'Shipments',1,'FR',300,rates,{'XPO':10},{'XPO':5},20,True,self.engine.find_rates)
        book=openpyxl.load_workbook(BytesIO(output))
        self.assertEqual(book['Shipments']['E2'].value,'=1+1')
        sheet=book['Staffel 1-4 EP']
        self.assertEqual(counts['bestemmingen'],2)
        self.assertEqual(sheet.max_row,9)
        for row,qty in zip(list(sheet.values)[1:5],range(1,5)):
            self.assertEqual(row[1],'01234')
            self.assertEqual(row[2],qty)
            self.assertEqual(row[3],qty*300)
            self.assertEqual(row[4],round(qty*0.4,2))
        headers=[cell.value for cell in sheet[1]]
        self.assertFalse(any('Drost' in header for header in headers))
        self.assertEqual(sheet.max_column,8)
        self.assertEqual(sheet.cell(5,headers.index('XPO - Totale inkoop EUR')+1).value,460)
        self.assertEqual(sheet.cell(5,headers.index('XPO - Verkoop EUR')+1).value,552)
        with self.assertRaises(ValueError):
            add_pallet_scale(self.workbook(),'Shipments',1,'FR',0,rates,{}, {},0,False,self.engine.find_rates)
