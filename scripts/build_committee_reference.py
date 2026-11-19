#!/usr/bin/env python3
"""Build editable committee masters from measured source-PDF geometry.

Each composition cites a specific source page. Business text/data are supplied
only by deckspec; the reusable PPTX contains labels and zero chart data.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_MARK
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

ROOT=Path(__file__).resolve().parents[1]
BG,INK,MUTED,BLUE,YELLOW,GREEN='E6EDF1','414C53','737982','2775B6','FFCC00','029154'


def dump(path,data):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n')


def build(source):
    p=Presentation();p.slide_width=Pt(720);p.slide_height=Pt(405)
    p.core_properties.title='Комитет — макеты по исходным слайдам'
    p.core_properties.author='';p.core_properties.last_modified_by=''
    p.core_properties.keywords='committee-source-layouts'
    maps={};provenance=[]
    def slide():
        s=p.slides.add_slide(p.slide_layouts[6]);s.background.fill.solid()
        s.background.fill.fore_color.rgb=RGBColor.from_string(BG);return s
    def text(s,name,value,box,size=10,bold=False,color=INK,cond=False):
        x,y,w,h=box;sh=s.shapes.add_textbox(Pt(x),Pt(y),Pt(w),Pt(h));sh.name=name
        tf=sh.text_frame;tf.clear();tf.word_wrap=True
        tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
        par=tf.paragraphs[0];par.space_after=Pt(0);par.line_spacing=1.1
        r=par.add_run();r.text=value;r.font.name='YS Text Cond' if cond else 'YS Text'
        r.font.size=Pt(size);r.font.bold=bold;r.font.italic=cond;r.font.color.rgb=RGBColor.from_string(color)
        return sh
    def panel(s,name,box,fill='FFFFFF',radius=14):
        x,y,w,h=box;sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Pt(x),Pt(y),Pt(w),Pt(h));sh.name=name
        sh.fill.solid();sh.fill.fore_color.rgb=RGBColor.from_string(fill);sh.line.fill.background()
        sh.adjustments[0]=min(.5,radius/min(w,h))
        style=sh._element.find(qn('p:style'))
        if style is not None:sh._element.remove(style)
        return sh
    def slot(sh,optional=False,repeat_names=None):
        r=sh.text_frame.paragraphs[0].runs[0]
        font='YS Text Cond Bold Italic' if r.font.italic else 'YS Text Bold' if r.font.bold else 'YS Text'
        d={'match_name':sh.name,'font':font,'size_pt':r.font.size.pt,'box_w_in':sh.width/914400,'box_h_in':sh.height/914400,'optional':optional,'clear_if_empty':optional,'clear_soft_breaks':True}
        if repeat_names:d.pop('match_name');d.update(repeat=True,match_names=repeat_names)
        return d
    def heading(s,value='ЗАГОЛОВОК СЛАЙДА',box=(28,14,664,31)):
        return text(s,'heading',value,box,21.6,True,cond=True)
    def add(s,family,slots,page,doc,chart_slots=None):
        anchor=f'committee:{family}:source';s.shapes[0].name=anchor
        for d in slots.values():
            if d.get('match_name')=='heading':d['match_name']=anchor
        v={'name':'source','_hint':len(p.slides)-1,'anchor':{'has_shapes':[anchor]},'slots':slots}
        if chart_slots:v['chart_slots']=chart_slots
        spec={'id':f'committee_{family}','schema_version':1,'template':'Комитет','family':family,'kind':'content_fill','strict_slots':True,'_doc':doc,'source_page':page,'variants':[v]}
        maps[family]=spec;provenance.append({'family':family,'variant':'source','source_pdf_page':page,'geometry':'measured from source PDF; no adaptive recentering or added card grid'})
    def table(s,name,box,cols,rows=6,size=8):
        x,y,w,h=box;fr=s.shapes.add_table(rows,len(cols),Pt(x),Pt(y),Pt(w),Pt(h));fr.name=name
        for col,cw in zip(fr.table.columns,cols):col.width=Pt(cw)
        for ri,row in enumerate(fr.table.rows):
            row.height=Pt(h/rows)
            for ci,cell in enumerate(row.cells):
                cell.fill.solid();cell.fill.fore_color.rgb=RGBColor.from_string('FFFFFF')
                cell.margin_left=Pt(8);cell.margin_right=Pt(8);cell.margin_top=Pt(4);cell.margin_bottom=Pt(4)
                cell.vertical_anchor=MSO_ANCHOR.TOP
                pr=cell._tc.get_or_add_tcPr()
                for side in ('L','R','T','B'):
                    ln=OxmlElement('a:ln'+side);ln.append(OxmlElement('a:noFill'));pr.append(ln)
                par=cell.text_frame.paragraphs[0];par.space_after=Pt(0);par.line_spacing=1.05
                r=par.add_run();r.text='Заголовок' if ri==0 else '—';r.font.name='YS Text';r.font.size=Pt(size)
                r.font.bold=ri==0 or ci==0;r.font.color.rgb=RGBColor.from_string(MUTED if ri==0 else INK)
        return fr
    def tslot(fr,rows,cols,size=8):
        return {'kind':'table','match_name':fr.name,'trim':True,'stretch':False,'stretch_font':False,'max_rows':rows,'max_cols':cols,'exact_cols':cols,'cell_font_pt':size,'fit_table':True,'divider_mode':'groups' if fr.name=='roadmap-table' else 'rows'}
    def chart(s,name,box,kind=XL_CHART_TYPE.COLUMN_CLUSTERED,nseries=2,colors=None,legend=True,labels=False):
        x,y,w,h=box;d=CategoryChartData();d.categories=['Период 1','Период 2','Период 3']
        for i in range(nseries):d.add_series('Ряд '+str(i+1),[0,0,0])
        fr=s.shapes.add_chart(kind,Pt(x),Pt(y),Pt(w),Pt(h),d);fr.name=name;c=fr.chart
        c.font.name='YS Text';c.font.size=Pt(7);c.has_legend=legend
        if legend:c.legend.position=XL_LEGEND_POSITION.BOTTOM;c.legend.include_in_layout=False
        c.category_axis.tick_labels.font.size=Pt(6.5);c.value_axis.tick_labels.font.size=Pt(6.5)
        for ax in (c.category_axis,c.value_axis):
            ax.tick_labels.font.name='YS Text';ax.tick_labels.font.bold=False;ax.tick_labels.font.color.rgb=RGBColor.from_string(MUTED)
            ax.major_tick_mark=XL_TICK_MARK.NONE;ax.minor_tick_mark=XL_TICK_MARK.NONE
            ax.format.line.fill.background()
        c.value_axis.has_major_gridlines=False;c.value_axis.minimum_scale=0
        for ser,col in zip(c.series,colors or ['B4B9BC',YELLOW,'D8DCDF',GREEN]):
            if kind==XL_CHART_TYPE.LINE:
                ser.format.line.color.rgb=RGBColor.from_string(col);ser.format.line.width=Pt(1.3)
            else:
                ser.format.fill.solid();ser.format.fill.fore_color.rgb=RGBColor.from_string(col);ser.format.line.fill.background()
        if kind in (XL_CHART_TYPE.COLUMN_CLUSTERED,XL_CHART_TYPE.COLUMN_STACKED):
            c.plots[0].gap_width=65
        if kind==XL_CHART_TYPE.COLUMN_STACKED:c.plots[0].overlap=100
        if labels:
            pl=c.plots[0];pl.has_data_labels=True;pl.data_labels.show_value=True
            pl.data_labels.position=XL_LABEL_POSITION.ABOVE if kind==XL_CHART_TYPE.LINE else XL_LABEL_POSITION.OUTSIDE_END
            pl.data_labels.font.size=Pt(7);pl.data_labels.font.color.rgb=RGBColor.from_string(INK)
        sp=OxmlElement('c:spPr');sp.append(OxmlElement('a:noFill'));ln=OxmlElement('a:ln');ln.append(OxmlElement('a:noFill'));sp.append(ln);c._chartSpace.insert_element_before(sp,'c:txPr','c:externalData')
        return fr

    # PDF page 1: the actual simple title, no invented white ribbon.
    s=slide();h=text(s,'heading','НАЗВАНИЕ ПРЕЗЕНТАЦИИ',(36,96.8,650,85),61.02,True,cond=True)
    date=text(s,'date','Дата комитета',(40.7,277.2,640,25),14.99,True,MUTED)
    add(s,'cover',{'title':slot(h),'date':slot(date)},1,'Обложка по PDF стр.1; палитра финальной дебиторки, без синего акцента.')
    s=slide();num=text(s,'heading','01',(38.5,44.6,200,90),64.85,True,INK,True)
    h=text(s,'section','НАЗВАНИЕ РАЗДЕЛА',(36,162,650,65),48.81,True,cond=True)
    add(s,'divider',{'number':slot(num),'title':slot(h)},4,'Разделитель по PDF стр.4: номер и название раздела слева; тёмный текст в палитре дебиторки.')
    s=slide();h=heading(s);sub=text(s,'subtitle','Период, сценарий и база сравнения',(32,46,656,15),9,color=MUTED)
    labels=[];values=[];baselines=[];bodies=[];focus=[]
    for i,x in enumerate([28.1,253.1,478.1],1):
        panel(s,f'panel-{i}',(x,64.7,213.8,188.9),radius=19)
        labels.append(text(s,f'label-{i}','Название метрики',(x+18,81,177.8,19),11.15,True))
        values.append(text(s,f'value-{i}','—',(x+18,107,177.8,29),26.5,True,GREEN,True))
        baselines.append(text(s,f'baseline-{i}','План / сравнение',(x+18,139,177.8,15),8.7,color=MUTED))
        bodies.append(text(s,f'body-{i}','Контекст и объяснение отклонения',(x+18,165,177.8,46),10.2))
        focus.append(text(s,f'focus-{i}','Фокус на следующий период',(x+18,225,177.8,25),9.6,True))
    fr=table(s,'summary-table',(29,270,663,112),[108,255,300],4,9.2)
    slots={'heading':slot(h),'subtitle':slot(sub),'table':tslot(fr,4,3,9.2)}
    for name,arr in [('labels',labels),('values',values),('baselines',baselines),('bodies',bodies),('focus',focus)]:slots[name]=slot(arr[0],repeat_names=[v.name for v in arr])
    add(s,'summary',slots,2,'Саммари по PDF стр.2: три содержательных KPI-блока с контекстом и фокусом; таблица направлений снизу.')
    # Page 8: chart and full-height substantive commentary, exact panel ratio.
    s=slide();h=heading(s,box=(34,14,652,32));panel(s,'chart-panel',(34,76.7,417.3,284.6),radius=28);panel(s,'note-panel',(459.2,76.7,226.8,284.6),radius=28)
    ct=text(s,'chart-title','Динамика показателя и сравнение с планом',(52,91,378,20),10,True)
    cf=chart(s,'main-chart',(48,121,390,223),nseries=2,labels=True)
    nh=text(s,'note-heading','Фокус следующего периода',(478,96,190,42),11,True)
    nb=text(s,'note-body','Действия, основанные на данных',(479,155,188,184),10.5)
    add(s,'chart_note',{'heading':slot(h),'chart_title':slot(ct),'note_heading':slot(nh),'note_body':slot(nb)},8,'График и рабочий план по PDF стр.8; пропорции панелей и плотность исходника.',{'chart':{'match_name':cf.name}})
    # Page 29: actual analytical dashboard with one large and three small charts.
    s=slide();h=heading(s,box=(37.5,10,660,34))
    for name,box in [('left',(37.5,58.1,358.3,296.6)),('top',(406.4,58.1,291.1,142.5)),('bottom',(406.4,210.7,291.1,144))]:panel(s,'panel-'+name,box,radius=15)
    slots={'heading':slot(h)};charts={}
    texts=[('left_title',(52,68,329,30),11,True),('left_subtitle',(52,101,329,14),8.5,False),('top_title',(418,68,266,18),10,True),('top_subtitle',(418,87,266,14),8,False),('bottom_title',(418,220,266,18),10,True),('bottom_subtitle',(418,240,266,14),8,False)]
    for name,box,size,bold in texts:slots[name]=slot(text(s,name,'Вывод и определение метрики',box,size,bold))
    boxes=[('main',(49,125,335,218),XL_CHART_TYPE.COLUMN_STACKED,3),('trend',(415,107,277,83),XL_CHART_TYPE.LINE,2),('step1',(415,265,130,79),XL_CHART_TYPE.LINE,1),('step2',(560,265,130,79),XL_CHART_TYPE.LINE,1)]
    for name,box,kind,n in boxes:
        fr=chart(s,name,box,kind,n,legend=name in ('main','trend'),labels=kind==XL_CHART_TYPE.LINE)
        charts[name]={'match_name':fr.name}
    add(s,'dashboard',slots,29,'Аналитический слайд по PDF стр.29: большой накопительный график и три компактных графика с подписями.',charts)
    # Page 25: actual ranked causes and an actionable strip at the bottom.
    s=slide();h=heading(s,box=(24,21,672,33));sub=text(s,'subtitle','Масштаб проблемы и база расчёта',(25,63,670,15),9,color=MUTED)
    panel(s,'ranking-panel',(24.1,89.2,671.8,291),radius=18)
    ch=text(s,'chart-title','ТОП-5 ПРИЧИН',(40,105,635,30),10.5,True)
    fr=chart(s,'ranking',(40,144,635,164),XL_CHART_TYPE.BAR_CLUSTERED,1,['AEAEAE'],False,True)
    panel(s,'callout-panel',(39.1,323.2,641.8,43.5),'FFF3CC',10)
    cl=text(s,'callout-label','ДАЛЬШЕ',(52,336,76,18),12,True,GREEN,True)
    cb=text(s,'callout','Действие и ожидаемое изменение',(133,332,531,30),9)
    add(s,'priorities',{'heading':slot(h),'subtitle':slot(sub),'chart_title':slot(ch),'callout_label':slot(cl),'callout':slot(cb)},25,'Приоритизация по PDF стр.25: горизонтальные причины с масштабом и компактный следующий шаг.',{'chart':{'match_name':fr.name}})
    # Page 30: dense roadmap table and genuine narrow next-steps column.
    s=slide();h=heading(s,box=(24,14,672,32));panel(s,'table-panel',(28.8,80.3,427.6,272.6),radius=8);panel(s,'next-panel',(465.9,80.3,174.1,272.6),radius=8)
    fr=table(s,'roadmap-table',(40.8,92,403.6,248),[85,143,64,111.6],9,8.4)
    nh=text(s,'next-heading','ДАЛЬНЕЙШИЕ ШАГИ',(482,98,142,20),9.5,True)
    nb=text(s,'next-body','Зависимости и следующие решения',(482,139,142,195),9.8)
    add(s,'roadmap',{'heading':slot(h),'table':tslot(fr,9,4,8.4),'next_heading':slot(nh),'next_body':slot(nb)},30,'Бэклог по PDF стр.30: до восьми проектов с увеличенными отступами, разделителями направлений и отдельной колонкой зависимостей.')
    folder=ROOT/'layouts/committee'
    for file in folder.glob('*.json'):file.unlink()
    for family,spec in maps.items():dump(folder/(family+'.json'),spec)
    path=ROOT/'Templates/committee/committee-template-v1.pptx';path.parent.mkdir(parents=True,exist_ok=True);p.save(path)
    dump(folder/'_canvas.json',{'slide_w_in':10,'slide_h_in':5.625,'margin_left_in':.33,'margin_right_in':.33,'heading':{'left_in':.39,'top_in':.19,'w_in':9.22,'h_in':.45},'content_band':{'left_in':.33,'right_in':9.67,'top_in':.8,'bottom_in':5.3,'w_in':9.34,'h_in':4.5},'gutter_in':.14})
    dump(folder/'brand-tokens.json',{'name':'Комитет','colors':{'background':BG,'panel':'FFFFFF','text':INK,'muted':MUTED,'accent':YELLOW,'positive':GREEN,'blue':BLUE},'fonts':{'heading':'YS Text Cond Bold Italic','body':'YS Text'},'reference':{'filename':source.name,'sha256':hashlib.sha256(source.read_bytes()).hexdigest()},'source_pages':[1,2,4,8,25,29,30]})
    dump(folder/'custom-profile.json',{'schema_version':1,'fonts':{'files':{'YS Text':'YS Text-Regular.ttf','YS Text Bold':'YS Text-Bold.ttf','YS Text Cond Bold Italic':'YS Text Cond-Bold Italic.ttf'},'fallback':'YS Text','resolve':[{'family_contains':'Cond','style':'YS Text Cond Bold Italic'},{'weight_min':700,'style':'YS Text Bold'},{'style':'YS Text'}],'bold_family':{'YS Text':'YS Text Bold'}},'layouts':{'default':'Blank','default_bg':[230,237,241],'by_bg':[]},'calibration':{'title_dy_px':0},'canvas':{'px':[720,405]},'a15':{'photo_panel':{'top_max':0,'bottom_min_from_h':0}},'margins':{'left':24,'right':24,'top':14,'bottom':18}})
    dump(ROOT/'Templates/committee/provenance.json',{'source_filename':source.name,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'source_modified':False,'copied_resources':'measured geometry and styles; native editable objects rebuilt','contains_business_content':False,'variants':provenance})
    mp=ROOT/'Templates/manifest.json';m=json.loads(mp.read_text());blob=path.read_bytes();m['templates']['committee']={'display_name':'Комитет','local_only':True,'template':{'name':path.name,'size':len(blob),'md5':hashlib.md5(blob).hexdigest(),'sha256':hashlib.sha256(blob).hexdigest()},'fonts':{'system_required':['YS Text-Regular.ttf','YS Text-Bold.ttf','YS Text Cond-Bold Italic.ttf']}};dump(mp,m)
    print('Built',len(p.slides),'source-based masters')


if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('source',type=Path);build(ap.parse_args().source)
