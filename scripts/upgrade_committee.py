#!/usr/bin/env python3
"""Add a clean chart-and-actions master based on the final committee slide.

The reference is hashed for provenance. Its business data and notes are never
copied. Run after build_committee_reference.py when rebuilding the original pack.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
BG, INK, MUTED, GREEN, YELLOW, BLUE = 'E6EDF1', '414C53', '737982', '029154', 'FFCC00', '2775B6'


def dump(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n')


def build(reference: Path):
    path = ROOT / 'Templates/committee/committee-template-v1.pptx'
    p = Presentation(path)
    if any(s.name == 'committee:chart_table:actions' for slide in p.slides for s in slide.shapes):
        raise ValueError('chart_table already exists; rebuild the base pack before upgrading')
    s = p.slides.add_slide(p.slide_layouts[6])
    s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor.from_string(BG)
    def txt(name, value, x,y,w,h,size,bold=False,italic=False,color=INK):
        sh = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); sh.name=name
        tf=sh.text_frame; tf.clear(); tf.word_wrap=True
        tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
        par=tf.paragraphs[0]; par.space_after=Pt(0); par.line_spacing=1.05
        r=par.add_run(); r.text=value; r.font.name='YS Text'; r.font.size=Pt(size)
        r.font.bold=bold; r.font.italic=italic; r.font.color.rgb=RGBColor.from_string(color)
        return sh
    def panel(name,x,y,w,h):
        sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h))
        sh.name=name; sh.fill.solid(); sh.fill.fore_color.rgb=RGBColor(255,255,255); sh.line.fill.background()
        style=sh._element.find(qn('p:style'))
        if style is not None:sh._element.remove(style)
        sh.adjustments[0]=.065
    title=txt('committee:chart_table:actions','НАЗВАНИЕ ПОКАЗАТЕЛЯ',.375,.177,9.25,.385,22.5,True,True)
    panel('chart-panel',.375,.656,9.25,2.083)
    unit=txt('chart-unit','Единица',.547,.739,1.5,.16,8.25,color=MUTED)
    data=CategoryChartData(); data.categories=['Период 1','Период 2','Период 3']
    for name in ['Компонент 1','Компонент 2','Доля, %']: data.add_series(name,[0,0,0])
    cf=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(.53), Inches(.88), Inches(8.94), Inches(1.76),data)
    cf.name='committee-combo'; chart=cf.chart
    chart.font.name='YS Text';chart.font.size=Pt(7.5)
    chart.has_legend=True; chart.legend.position=XL_LEGEND_POSITION.TOP;chart.legend.include_in_layout=False
    chart.category_axis.tick_labels.font.size=Pt(6.75)
    chart.value_axis.tick_labels.font.size=Pt(6.75);chart.value_axis.minimum_scale=0
    chartsp=OxmlElement('c:spPr');chartsp.append(OxmlElement('a:noFill'))
    chartln=OxmlElement('a:ln');chartln.append(OxmlElement('a:noFill'));chartsp.append(chartln)
    chart._chartSpace.insert_element_before(chartsp,'c:txPr','c:externalData')
    for ax in (chart.category_axis,chart.value_axis):
        ax.major_tick_mark=XL_TICK_MARK.NONE;ax.minor_tick_mark=XL_TICK_MARK.NONE
        ax.tick_labels.font.name='YS Text';ax.tick_labels.font.bold=False
        ax.tick_labels.font.color.rgb=RGBColor.from_string(MUTED)
    chart.category_axis.format.line.color.rgb=RGBColor.from_string('DCE3E7')
    chart.category_axis.format.line.width=Pt(.5)
    chart.value_axis.format.line.fill.background()
    chart.value_axis.has_major_gridlines=True
    chart.value_axis.major_gridlines.format.line.color.rgb=RGBColor.from_string('EDF1F3')
    chart.value_axis.major_gridlines.format.line.width=Pt(.5)
    plot=chart.plots[0];plot.gap_width=55;plot.overlap=100
    for ser,col in zip(chart.series,[GREEN,YELLOW,BLUE]):
        ser.format.fill.solid();ser.format.fill.fore_color.rgb=RGBColor.from_string(col)
        ser.format.line.fill.background()
    bar=chart._chartSpace.find('.//'+qn('c:barChart'))
    third=bar.findall(qn('c:ser'))[-1];bar.remove(third)
    line=OxmlElement('c:lineChart');group=OxmlElement('c:grouping');group.set('val','standard');line.append(group)
    line.append(third)
    sp=third.find(qn('c:spPr'))
    for child in list(sp):sp.remove(child)
    ln=OxmlElement('a:ln');ln.set('w','19050');solid=OxmlElement('a:solidFill');rgb=OxmlElement('a:srgbClr');rgb.set('val',BLUE);solid.append(rgb);ln.append(solid);sp.append(ln)
    axids=[x.get('val') for x in bar.findall(qn('c:axId'))]; secondary='999100'
    for axid in (axids[0],secondary):
        node=OxmlElement('c:axId');node.set('val',axid);line.append(node)
    area=chart._chartSpace.find('.//'+qn('c:plotArea'));bar.addnext(line)
    axis=copy.deepcopy(area.find(qn('c:valAx')))
    axis.find(qn('c:axId')).set('val',secondary);axis.find(qn('c:axPos')).set('val','r')
    axis.find(qn('c:crossAx')).set('val',axids[0])
    for child in list(axis):
        if child.tag in (qn('c:majorGridlines'),qn('c:numFmt')):axis.remove(child)
    scale=axis.find(qn('c:scaling'))
    for key,val in [('min','0'),('max','100')]:
        node=OxmlElement('c:'+key);node.set('val',val);scale.append(node)
    axis.find(qn('c:tickLblPos')).set('val','none')
    area.append(axis)
    # Percentage data labels remain editable, and no old point labels survive.
    lp=chart.plots[1];lp.has_data_labels=True
    chart.series[2].marker.style=XL_MARKER_STYLE.CIRCLE;chart.series[2].marker.size=3
    lp.data_labels.show_value=True;lp.data_labels.number_format='0"%"'
    lp.data_labels.position=XL_LABEL_POSITION.ABOVE;lp.data_labels.font.size=Pt(6.75)
    lp.data_labels.font.color.rgb=RGBColor.from_string(BLUE)
    panel('table-panel',.375,2.844,9.25,2.604)
    tf=s.shapes.add_table(8,4, Inches(.542), Inches(2.98), Inches(8.916), Inches(2.34));tf.name='committee-actions'
    for col,w in zip(tf.table.columns,[1.875,.86,4.95,1.231]):col.width=Inches(w)
    for ri,row in enumerate(tf.table.rows):
        row.height=Inches(.2925)
        for ci,cell in enumerate(row.cells):
            cell.fill.solid();cell.fill.fore_color.rgb=RGBColor(255,255,255)
            cell.margin_left=0;cell.margin_right=Inches(.07);cell.margin_top=Inches(.05);cell.margin_bottom=Inches(.035)
            cp=cell._tc.get_or_add_tcPr()
            for side in ('L','R','T','B'):
                node=OxmlElement('a:ln'+side);node.append(OxmlElement('a:noFill'));cp.append(node)
            par=cell.text_frame.paragraphs[0];par.space_after=Pt(0);par.line_spacing=1.0
            r=par.add_run();r.text=['ПРИЧИНА','ДОЛЯ','ЧТО ДЕЛАЕМ','СРОК'][ci] if ri==0 else '—'
            r.font.name='YS Text';r.font.size=Pt(8.0 if ri==0 else 9.0)
            r.font.bold=ri==0 or ci in (0,1);r.font.italic=ri==0
            r.font.color.rgb=RGBColor.from_string(MUTED if ri==0 else GREEN if ci==1 else INK)
    def slot(sh,size,w,h):return {'match_name':sh.name,'font':'YS Text','size_pt':size,'box_w_in':w,'box_h_in':h}
    spec={'id':'committee_chart_table','schema_version':1,'template':'Комитет','family':'chart_table','kind':'content_fill','strict_slots':True,
          '_doc':'Динамика сверху, причины/доли/действия/сроки снизу. Два накопительных ряда + процентная линия; нативные график и таблица.',
          'variants':[{'name':'actions','_hint':len(p.slides)-1,'anchor':{'has_shapes':[title.name]},
           'slots':{'heading':slot(title,22.5,9.25,.385),'unit':slot(unit,8.25,1.5,.16),
                    'table':{'kind':'table','match_name':tf.name,'trim':True,'stretch':False,'stretch_width':False,'stretch_font':False,'max_rows':8,'max_cols':4,'exact_cols':4}},
           'chart_slots':{'chart':{'match_name':cf.name,'combo_series':3}}}]}
    dump(ROOT/'layouts/committee/chart_table.json',spec)
    p.save(path);blob=path.read_bytes()
    mp=ROOT/'Templates/manifest.json';manifest=json.loads(mp.read_text());pin=manifest['templates']['committee']['template']
    pin.update(size=len(blob),md5=hashlib.md5(blob).hexdigest(),sha256=hashlib.sha256(blob).hexdigest());dump(mp,manifest)
    tp=ROOT/'layouts/committee/brand-tokens.json';tokens=json.loads(tp.read_text());tokens['colors'].update(background=BG,text=INK,accent=YELLOW,muted=MUTED,line=BLUE);dump(tp,tokens)
    cp=ROOT/'layouts/committee/custom-profile.json';profile=json.loads(cp.read_text());profile['layouts']['default_bg']=[230,237,241];dump(cp,profile)
    prov=ROOT/'Templates/committee/provenance.json';info=json.loads(prov.read_text())
    info['final_slide_reference']={'filename':reference.name,'sha256':hashlib.sha256(reference.read_bytes()).hexdigest(),'copied_resources':'geometry and style only; chart, table and labels regenerated without business data'}
    info['variants'].append({'family':'chart_table','variant':'actions','source':'final_slide_reference'})
    dump(prov,info)
    print(json.dumps({'variants':len(p.slides),'bytes':len(blob),'sha256':pin['sha256']}))


if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('reference',type=Path);build(ap.parse_args().reference)
