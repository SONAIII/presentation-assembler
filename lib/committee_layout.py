"""Balance committee compositions around the actual amount of text.

Resize and place existing donor elements only. Do not add content, reduce font
sizes or stretch sparse copy to the full content band.
"""
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
import fit


def apply(slide, family, variant):
    by_name = {s.name: s for s in slide.shapes}
    slots = variant["slots"]
    top, bottom = 1.5, 5.0
    center = (top + bottom) / 2

    def divider(name,x,y,w):
        sh=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,int(x),int(y),int(x+w),int(y))
        sh.name=name;sh.line.width=Pt(.55);sh.line.color.rgb=RGBColor.from_string('DCE3E7')
        style=sh._element.find(qn('p:style'))
        if style is not None:sh._element.remove(style)

    if variant.get('name') == 'source':
        # Source-PDF masters keep their measured geometry. Do not center sparse
        # content or invent taller cards. Catch overflowing native cells here.
        for meta in slots.values():
            if not meta.get('fit_table'):
                continue
            fr=by_name[meta['match_name']];tbl=fr.table
            for ri,row in enumerate(tbl.rows):
                for ci,cell in enumerate(row.cells):
                    size=meta['cell_font_pt'];font='YS Text Bold' if ri==0 or ci==0 else 'YS Text'
                    usable=(tbl.columns[ci].width-cell.margin_left-cell.margin_right)/914400
                    needed=fit.text_height_in(cell.text,font,size,usable,line_spacing=1.05)+(cell.margin_top+cell.margin_bottom)/914400
                    if needed>row.height/914400+.025:
                        raise ValueError(f'Таблица {fr.name}, строка {ri+1}: текст не вмещается при исходном кегле; сократи формулировку или выбери другой исходный макет')
            y=fr.top
            for ri,row in enumerate(tbl.rows):
                y+=row.height
                if ri==len(tbl.rows)-1:continue
                if meta.get('divider_mode')=='groups' and ri>0 and tbl.cell(ri,0).text==tbl.cell(ri+1,0).text:continue
                divider(f'{fr.name}:divider:{ri}',fr.left,y,fr.width)
        if family=='summary':
            for i in range(1,4):
                sh=by_name[f'focus-{i}'];divider(f'focus-divider-{i}',sh.left,sh.top-Pt(8),sh.width)
        if family in ('chart_note','roadmap'):
            a=by_name['note-heading' if family=='chart_note' else 'next-heading']
            divider('comment-divider',a.left,a.top+a.height+Pt(5),a.width)
        return {'layout':'committee','source_geometry_preserved':True}

    def shapes(key):
        spec = slots[key]
        return [by_name[n] for n in spec.get("match_names", [spec.get("match_name")])]

    def height(sh):
        tf = sh.text_frame
        run = next((r for p in tf.paragraphs for r in p.runs if r.text.strip()), None)
        if run is None:
            return 0
        name = "YS Text Bold" if run.font.bold else "YS Text"
        size = run.font.size.pt
        width = sh.width / 914400 - .04
        for p in tf.paragraphs:
            p.line_spacing = 1.15
        return sum(fit.text_height_in(line, name, size, width, line_spacing=1.15)
                   for line in tf.text.splitlines() if line.strip()) + .04

    def move(sh, y, h):
        sh.top, sh.height = Inches(y), Inches(h)

    def center_panel(panel, h):
        y = center - h / 2
        move(panel, y, h)
        return y

    if family == 'chart_table':
        frame=by_name[slots['table']['match_name']]
        table=frame.table
        heights=[]
        for ri,row in enumerate(table.rows):
            required=.29
            for ci,cell in enumerate(row.cells):
                if ci==1:
                    for par in cell.text_frame.paragraphs:par.alignment=PP_ALIGN.LEFT
                w=table.columns[ci].width/914400-.08
                size=8 if ri==0 else 9
                h=fit.text_height_in(cell.text,'YS Text Bold' if ri==0 or ci in (0,1) else 'YS Text',size,w,line_spacing=1.05)+.085
                required=max(required,h)
                if ri and ci==3 and cell.text.strip()=='Выполнено':
                    for p in cell.text_frame.paragraphs:
                        for r in p.runs:r.font.bold=True;r.font.color.rgb=RGBColor.from_string('029154')
            heights.append(required)
        if sum(heights)>2.34:
            raise ValueError('chart_table: текст таблицы не вмещается при читаемом кегле; сократи формулировки или раздели слайды')
        for row,h in zip(table.rows,heights):row.height=Inches(h)
        frame.height=Inches(sum(heights))
        by_name['table-panel'].height=frame.height+Inches(.27)
        divider('actions-header-divider',frame.left,frame.top+table.rows[0].height,frame.width)
        # Empty repeated cause/share cells mean more actions for the same cause.
        starts=[r for r in range(1,len(table.rows)) if table.cell(r,0).text.strip()]
        for k,start in enumerate(starts):
            end=(starts[k+1]-1) if k+1<len(starts) else len(table.rows)-1
            if end>start and all(not table.cell(r,c).text.strip() for r in range(start+1,end+1) for c in (0,1)):
                for c in (0,1):table.cell(start,c).merge(table.cell(end,c))
            if k+1<len(starts):
                y=frame.top+sum(row.height for row in list(table.rows)[:end+1])
                divider(f'actions-divider-{k}',frame.left,y,frame.width)
                for ci in range(4):
                    cp=table.cell(start if ci<2 else end,ci)._tc.get_or_add_tcPr()
                    old=cp.find(qn('a:lnB'))
                    if old is not None:cp.remove(old)
                    ln=OxmlElement('a:lnB');ln.set('w','6350');fill=OxmlElement('a:solidFill')
                    col=OxmlElement('a:srgbClr');col.set('val','DCE3E7');fill.append(col);ln.append(fill);cp.append(ln)
        return {'layout':'committee','native_table':True,'row_heights_in':heights}
    return {'layout':'committee','source_geometry_preserved':True}
