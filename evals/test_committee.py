# coding: utf-8
"""Offline checks for the source-based committee masters and native data path."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from pptx import Presentation
from lxml import etree
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import build_deck as B
import fetch_template as FT
import preflight as PF
import anchor

class CommitteePack(unittest.TestCase):
    def setUp(self):
        self.path=Path(FT.resolve_template('комитет'))
        self.prs=Presentation(self.path)
        self.demo=json.loads((ROOT/'examples/committee_smoke.json').read_text())
    def example(self,f):return copy.deepcopy(next(e for e in self.demo['slides'] if e['family']==f))
    def plan(self,e):return B.resolve_entry_plan(self.prs,e,'committee')
    def test_menu(self):
        self.assertEqual(FT.TDIR_ALIASES['комитет'],'committee')
        self.assertEqual(FT.TDIR_ALIASES['еда'],'eda')
        self.assertEqual(set(B.known_families('committee')),{'cover','divider','summary','chart_note','dashboard','priorities','roadmap','chart_table'})
    def test_unique_anchors_and_provenance(self):
        indices=[]
        for f in B.known_families('committee'):
            d=json.loads((ROOT/'layouts/committee'/f'{f}.json').read_text())
            indices.extend(anchor.resolve_entry(v,self.prs) for v in d['variants'])
        self.assertEqual(len(indices),len(set(indices)))
        self.assertEqual(set(indices),set(range(len(self.prs.slides))))
        prov=json.loads((ROOT/'Templates/committee/provenance.json').read_text())
        self.assertEqual({v['source_pdf_page'] for v in prov['variants'] if 'source_pdf_page' in v},{1,2,4,8,25,29,30})
        self.assertIn('final_slide_reference',prov)
    def test_clean_master(self):
        with zipfile.ZipFile(self.path) as z:
            self.assertFalse(any(n.startswith(('ppt/media/','ppt/notesSlides/')) for n in z.namelist()))
            for n in z.namelist():
                if n.endswith(('.xml','.rels')):self.assertFalse(etree.fromstring(z.read(n)).xpath('//*[@TargetMode="External"]'))
        for slide in self.prs.slides:
            for sh in slide.shapes:
                if sh.has_chart:
                    for ser in sh.chart.series:self.assertTrue(all(v==0 for v in ser.values))
    def test_pin_and_offline(self):
        pin=FT.load_manifest()['templates']['committee']['template'];self.assertTrue(FT._matches_pin(self.path,pin))
        with tempfile.TemporaryDirectory() as folder:
            bad=Path(folder)/'bad.pptx';bad.write_bytes(self.path.read_bytes()+b'changed')
            with self.assertRaises(SystemExit):FT._require_pin(str(bad),pin,'test')
        rep=PF.Report()
        with patch.object(FT,'_remote_meta') as remote:PF.check_template(rep,'committee',offline=False);remote.assert_not_called()
        self.assertFalse(rep.blockers)
    def test_geometry_and_condensed_font_preserved(self):
        plan=self.plan(self.example('summary'));src=self.prs.slides[plan['src']]
        boxes={s.name:(s.left,s.top,s.width,s.height) for s in src.shapes}
        idx,_=B.build_slide_on_clone(self.prs,plan,'committee');out=self.prs.slides[idx]
        self.assertEqual(boxes,{s.name:(s.left,s.top,s.width,s.height) for s in out.shapes if s.name in boxes})
        value=next(s for s in out.shapes if s.name=='value-1')
        self.assertEqual(value.text_frame.paragraphs[0].runs[0].font.name,'YS Text Cond')
    def test_entries_resolve(self):
        for e in self.demo['slides']:
            with self.subTest(family=e['family']):self.assertEqual(self.plan(e)['mode'],'FILL')
    def test_required_content_and_repeat_capacity(self):
        for f,k in [('cover','title'),('summary','bodies'),('chart_note','note_body'),('roadmap','table')]:
            e=self.example(f);e['slots'].pop(k)
            with self.assertRaises((B.DeckError,ValueError)):self.plan(e)
        e=self.example('summary');e['slots']['labels'].append('extra')
        with self.assertRaises(B.DeckError):self.plan(e)
    def test_lossy_tables_rejected(self):
        for value in ([['a','b'],['c']],[[None]*4],[['a']*4]*15,[['a','b'],['c','d']]):
            e=self.example('roadmap');e['slots']['table']=value
            with self.assertRaises(B.DeckError):self.plan(e)
    def test_bad_combo_rejected(self):
        for values in ([1,2],[float('nan')]*12,[101]*12,[-1]*12):
            e=self.example('chart_table');e['charts']['chart']['series']['Погашено, %']=values
            with self.assertRaises(B.DeckError):self.plan(e)
        e=self.example('chart_table');e['charts']['chart']['series'].pop('Погашено, %')
        with self.assertRaises(B.DeckError):self.plan(e)
    def test_missing_chart_rejected(self):
        e=self.example('dashboard');e['charts'].pop('trend')
        with self.assertRaises(B.DeckError):self.plan(e)
    def test_long_native_cells_not_clipped(self):
        for f in ('roadmap','chart_table'):
            e=self.example(f);e['slots']['table'][1][2]='Длинная формулировка для отдельного слайда. '*40
            with self.assertRaises(ValueError):B.build_slide_on_clone(self.prs,self.plan(e),'committee')
    def test_build_readback(self):
        with tempfile.TemporaryDirectory() as folder:
            out=Path(folder)/'demo.pptx'
            result=subprocess.run([sys.executable,str(ROOT/'scripts/build_deck.py'),str(ROOT/'examples/committee_smoke.json'),'-o',str(out)],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
            p=Presentation(out);self.assertEqual(len(p.slides),len(self.demo['slides']))
            for e,slide in zip(self.demo['slides'],p.slides):
                shapes={s.name:s for s in slide.shapes}
                v=json.loads((ROOT/'layouts/committee'/f"{e['family']}.json").read_text())['variants'][0]
                for name,want in e.get('charts',{}).items():
                    c=shapes[v['chart_slots'][name]['match_name']].chart
                    self.assertEqual([(s.name,list(s.values)) for s in c.series],list(want['series'].items()))
                    self.assertEqual([c.label.replace('\u00a0',' ') for c in c.plots[0].categories],want['categories'])
                if 'table' in e['slots']:
                    t=shapes[v['slots']['table']['match_name']].table
                    for ri,row in enumerate(e['slots']['table']):
                        for ci,value in enumerate(row):self.assertEqual(t.cell(ri,ci).text.replace('\u00a0',' '),value.rstrip('.'))
                self.assertIn('условных',slide.notes_slide.notes_text_frame.text)
    def test_duplicate_chart_guard(self):
        with tempfile.TemporaryDirectory() as folder:
            spec=Path(folder)/'bad.json';dest=Path(folder)/'bad.pptx'
            spec.write_text(json.dumps({'template':'committee','slides':[self.example('dashboard')]*2}))
            p=subprocess.run([sys.executable,str(ROOT/'scripts/build_deck.py'),str(spec),'-o',str(dest)],capture_output=True,text=True)
            self.assertNotEqual(p.returncode,0);self.assertIn('chart/OLE',p.stderr);self.assertFalse(dest.exists())
if __name__=='__main__':unittest.main()
