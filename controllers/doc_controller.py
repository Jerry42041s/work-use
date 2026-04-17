import base64
import io
import json
import re
import shutil
import subprocess
import tempfile
import os
import zipfile
import html as html_mod
from lxml import etree
from odoo import http
from odoo.http import request

# ─── DOCX XML 命名空間 ────────────────────────────────────────────────────────
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML = "http://www.w3.org/XML/1998/namespace"

# DOCX 段落樣式名稱 → HTML 標籤對應表（大小寫不敏感）
_HEADING_STYLES = {
    'heading1': 'h1', 'heading 1': 'h1',
    'heading2': 'h2', 'heading 2': 'h2',
    'heading3': 'h3', 'heading 3': 'h3',
    'heading4': 'h4', 'heading 4': 'h4',
    'heading5': 'h5', 'heading 5': 'h5',
    'heading6': 'h6', 'heading 6': 'h6',
}


def _docx_to_html_with_format(file_bytes):
    """使用 zipfile + lxml 直接解析 DOCX XML，保留字體大小、粗斜體、顏色、對齊等格式。
    無需安裝額外套件：zipfile 為 Python 標準函式庫，lxml 已隨 Odoo 安裝。
    """
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            doc_xml = zf.read('word/document.xml')
            try:
                styles_xml = zf.read('word/styles.xml')
            except KeyError:
                styles_xml = None
    except Exception as e:
        return f'<p>（無法解析 DOCX：{html_mod.escape(str(e))}）</p>'

    # ── 解析段落樣式的預設字體大小 ──────────────────────────────────────────
    style_default_sizes = {}   # styleId → font-size (pt)
    style_display_names = {}   # styleId → displayName (lower)
    if styles_xml:
        try:
            st = etree.fromstring(styles_xml)
            for style in st.findall(f'{{{_W}}}style'):
                sid = style.get(f'{{{_W}}}styleId') or ''
                name_el = style.find(f'{{{_W}}}name')
                display = (name_el.get(f'{{{_W}}}val') or '').lower() if name_el is not None else ''
                style_display_names[sid] = display
                sz = style.find(f'.//{{{_W}}}rPr/{{{_W}}}sz')
                if sz is not None:
                    val = sz.get(f'{{{_W}}}val')
                    if val and val.isdigit():
                        style_default_sizes[sid] = int(val) / 2  # half-point → pt
        except Exception:
            pass

    # ── 解析正文 ─────────────────────────────────────────────────────────────
    try:
        body = etree.fromstring(doc_xml).find(f'{{{_W}}}body')
    except Exception as e:
        return f'<p>（DOCX document.xml 解析失敗：{html_mod.escape(str(e))}）</p>'

    parts = []
    for child in body:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local == 'p':
            parts.append(_w_paragraph_to_html(child, style_default_sizes, style_display_names))
        elif local == 'tbl':
            parts.append(_w_table_to_html(child, style_default_sizes, style_display_names))
        # sectPr（節屬性）及其他 meta 略過

    return '\n'.join(parts) if parts else '<p></p>'


def _w_paragraph_to_html(p, style_sizes, style_names):
    """將 <w:p> 轉換為 HTML 段落或標題。"""
    pPr = p.find(f'{{{_W}}}pPr')
    style_id = ''
    align = ''

    if pPr is not None:
        ps = pPr.find(f'{{{_W}}}pStyle')
        if ps is not None:
            style_id = ps.get(f'{{{_W}}}val') or ''
        jc = pPr.find(f'{{{_W}}}jc')
        if jc is not None:
            jc_val = jc.get(f'{{{_W}}}val', '')
            align = {'center': 'center', 'right': 'right',
                     'both': 'justify', 'distribute': 'justify'}.get(jc_val, '')

    # 決定 HTML 標籤
    display_name = style_names.get(style_id, '').lower()
    tag = _HEADING_STYLES.get(display_name, 'p')

    # 處理 runs
    runs_html = _w_runs_to_html(p, style_sizes.get(style_id))

    content = ''.join(runs_html) or '<br>'
    style_attr = f' style="text-align:{align}"' if align else ''
    return f'<{tag}{style_attr}>{content}</{tag}>'


def _w_runs_to_html(p, default_font_size_pt):
    """將段落內的 <w:r> 轉換為帶 inline style 的 HTML 片段清單。"""
    runs = []
    for child in p:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if local == 'r':
            runs.append(_w_run_to_html(child, default_font_size_pt))
        elif local == 'hyperlink':
            # 超連結：遞迴處理內部 runs
            href = child.get(f'{{{_W}}}anchor') or ''
            inner = ''.join(_w_runs_to_html(child, default_font_size_pt))
            if href:
                runs.append(f'<a href="#{html_mod.escape(href)}">{inner}</a>')
            else:
                runs.append(inner)

    return runs


def _w_run_to_html(r, default_font_size_pt):
    """將單一 <w:r> 轉換為 HTML 字串，保留格式。"""
    rPr = r.find(f'{{{_W}}}rPr')

    # 收集文字（含 <w:t> 與 <w:br>）
    text_parts = []
    for child in r:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local == 't':
            # xml:space="preserve" → 保留空白
            preserve = child.get(f'{{{_XML}}}space') == 'preserve'
            t = child.text or ''
            text_parts.append(html_mod.escape(t) if not preserve else
                               html_mod.escape(t).replace(' ', '&nbsp;'))
        elif local == 'br':
            text_parts.append('<br>')

    if not text_parts:
        return ''

    text = ''.join(text_parts)

    if rPr is None:
        return f'<span style="font-size:{default_font_size_pt}pt">{text}</span>' \
            if default_font_size_pt else text

    # ── 解析格式屬性 ─────────────────────────────────────────────────────────
    css = []

    # 字體大小（half-points → pt）
    sz = rPr.find(f'{{{_W}}}sz')
    if sz is not None:
        val = sz.get(f'{{{_W}}}val')
        if val and val.isdigit():
            css.append(f'font-size: {int(val) // 2}pt')
    elif default_font_size_pt:
        css.append(f'font-size: {default_font_size_pt}pt')

    # 文字顏色（排除 "auto"）
    color = rPr.find(f'{{{_W}}}color')
    if color is not None:
        val = color.get(f'{{{_W}}}val', 'auto')
        if val and val.lower() not in ('auto', 'ffffff'):
            css.append(f'color: #{val}')

    # 字體名稱
    fonts = rPr.find(f'{{{_W}}}rFonts')
    if fonts is not None:
        ff = fonts.get(f'{{{_W}}}ascii') or fonts.get(f'{{{_W}}}eastAsia') or ''
        if ff:
            css.append(f"font-family: '{ff}'")

    # 行內格式（粗體、斜體、底線）
    bold = rPr.find(f'{{{_W}}}b') is not None
    italic = rPr.find(f'{{{_W}}}i') is not None
    underline = rPr.find(f'{{{_W}}}u') is not None

    # 套用格式
    style_attr = f' style="{"; ".join(css)}"' if css else ''
    result = f'<span{style_attr}>{text}</span>' if css else text
    if bold:
        result = f'<strong>{result}</strong>'
    if italic:
        result = f'<em>{result}</em>'
    if underline:
        result = f'<u>{result}</u>'

    return result


def _w_table_to_html(tbl, style_sizes, style_names):
    """將 <w:tbl> 轉換為 HTML 表格（使用 <tbody>，全部 <td> 以相容 Odoo html_editor）。"""
    rows = tbl.findall(f'.//{{{_W}}}tr')
    if not rows:
        return ''
    rows_html = []
    for tr in rows:
        cells = tr.findall(f'{{{_W}}}tc')
        cells_html = []
        for tc in cells:
            colspan = 1
            rowspan = 1
            tcPr = tc.find(f'{{{_W}}}tcPr')
            if tcPr is not None:
                gridSpan = tcPr.find(f'{{{_W}}}gridSpan')
                if gridSpan is not None:
                    val = gridSpan.get(f'{{{_W}}}val')
                    if val and val.isdigit():
                        colspan = int(val)
                vMerge = tcPr.find(f'{{{_W}}}vMerge')
                if vMerge is not None and vMerge.get(f'{{{_W}}}val') == 'restart':
                    rowspan = 2  # 簡化處理，精確 rowspan 需追蹤合併狀態
            span_attrs = ''
            if colspan > 1:
                span_attrs += f' colspan="{colspan}"'
            if rowspan > 1:
                span_attrs += f' rowspan="{rowspan}"'
            # 儲存格內可能有多個段落（<p> 包裝讓 ev.target = p，避免 resize 觸發）
            inner_parts = []
            for p in tc.findall(f'{{{_W}}}p'):
                inner_parts.append(_w_paragraph_to_html(p, style_sizes, style_names))
            inner = ''.join(inner_parts) or '<p>&nbsp;</p>'
            cells_html.append(f'<td{span_attrs}>{inner}</td>')
        rows_html.append(f'<tr>{"".join(cells_html)}</tr>')
    return (
        '<table style="width:100%">'
        '<tbody>'
        + ''.join(rows_html)
        + '</tbody>'
        + '</table>'
    )


def _odt_to_html(file_bytes):
    """使用 odfpy 將 ODT 轉換為 HTML，保留對齊、字體大小、行內格式等樣式。

    odfpy 的節點型態皆為 Element，需透過 node.qname[1]（local tag name）
    來識別 p / h / table 等節點。樣式屬性需用 (namespace_uri, localname) tuple 取得。
    """
    try:
        from odf.opendocument import load as odf_load
        from odf.namespaces import FONS, STYLENS, TEXTNS, XLINKNS
        from odf.element import Node as OdfNode
    except ImportError:
        return '<p>（無法解析 ODT：odfpy 未安裝）</p>'

    import html as html_mod

    doc = odf_load(io.BytesIO(file_bytes))
    parts = []

    def _local(node):
        """取得節點的 local tag name（如 'p', 'h', 'table'）。"""
        if hasattr(node, 'qname') and node.qname:
            return node.qname[1].lower()
        return ''

    # ── 建立 stylename → CSS dict 的對應表 ──
    def _extract_css_from_style(style_node):
        """從 style 節點提取 CSS 屬性字典（段落屬性 + 文字屬性）。"""
        css = {}
        for child in style_node.childNodes:
            if child.nodeType != OdfNode.ELEMENT_NODE:
                continue
            ln = _local(child)
            if ln == 'paragraph-properties':
                align = child.attributes.get((FONS, 'text-align'))
                if align:
                    if align == 'start':
                        align = 'left'
                    elif align == 'end':
                        align = 'right'
                    if align in ('center', 'right', 'justify'):
                        css['text-align'] = align
            elif ln == 'text-properties':
                fs = child.attributes.get((FONS, 'font-size'))
                if fs:
                    css['font-size'] = fs
                fw = child.attributes.get((FONS, 'font-weight'))
                if fw == 'bold':
                    css['font-weight'] = 'bold'
                fi = child.attributes.get((FONS, 'font-style'))
                if fi == 'italic':
                    css['font-style'] = 'italic'
                fc = child.attributes.get((FONS, 'color'))
                if fc and fc != 'transparent':
                    css['color'] = fc
        return css

    # 收集所有樣式（named styles → automatic styles 順序，自動樣式可覆蓋）
    style_map = {}
    for source in [getattr(doc, 'styles', None), doc.automaticstyles]:
        if source is None:
            continue
        for node in source.childNodes:
            if node.nodeType != OdfNode.ELEMENT_NODE:
                continue
            name = node.attributes.get((STYLENS, 'name'))
            if name:
                css = _extract_css_from_style(node)
                # 若有 parent-style-name，先繼承父樣式
                parent = node.attributes.get((STYLENS, 'parent-style-name'))
                if parent and parent in style_map:
                    merged = dict(style_map[parent])
                    merged.update(css)
                    style_map[name] = merged
                elif css:
                    style_map[name] = css

    def _para_style_attr(node):
        """從段落/標題節點取得 style="" 屬性字串（僅段落層級屬性）。"""
        sname = node.attributes.get((TEXTNS, 'style-name'))
        if not sname:
            return ''
        css = style_map.get(sname, {})
        if not css:
            return ''
        return ' style="' + '; '.join(f'{k}: {v}' for k, v in css.items()) + '"'

    def _inline_html(node):
        """遞迴將 ODT inline 節點轉換為 HTML，保留 text:span 的字體大小等樣式。"""
        result = []
        for child in node.childNodes:
            if child.nodeType == OdfNode.TEXT_NODE:
                # 純文字節點：直接 escape
                result.append(html_mod.escape(child.data or ''))
            elif child.nodeType == OdfNode.ELEMENT_NODE:
                if not hasattr(child, 'qname') or not child.qname:
                    continue
                ln = child.qname[1].lower()

                if ln == 'span':
                    # text:span — 取得樣式並遞迴處理內容
                    sname = child.attributes.get((TEXTNS, 'style-name'))
                    inner = _inline_html(child)
                    if sname and sname in style_map:
                        # 行內元素：排除 text-align（屬於段落屬性）
                        css = {k: v for k, v in style_map[sname].items()
                               if k != 'text-align'}
                        css_str = '; '.join(f'{k}: {v}' for k, v in css.items())
                        if css_str:
                            result.append(f'<span style="{css_str}">{inner}</span>')
                        else:
                            result.append(inner)
                    else:
                        result.append(inner)
                elif ln == 'line-break':
                    result.append('<br>')
                elif ln == 's':
                    # text:s — 多個空格
                    count_str = child.getAttribute('c') or '1'
                    try:
                        count = max(1, int(count_str))
                    except (ValueError, TypeError):
                        count = 1
                    result.append('&nbsp;' * count)
                elif ln == 'tab':
                    result.append('&nbsp;&nbsp;&nbsp;&nbsp;')
                elif ln == 'a':
                    # text:a — 超連結
                    inner = _inline_html(child)
                    href = child.attributes.get((XLINKNS, 'href'), '#')
                    result.append(
                        f'<a href="{html_mod.escape(href)}">{inner}</a>'
                    )
                else:
                    # 其他行內節點（text:ruby 等）→ 遞迴取文字
                    result.append(_inline_html(child))
        return ''.join(result)

    def _cell_paragraphs(cell_node):
        """將 ODT table cell 內的段落轉換為帶樣式的 HTML（保留 font-size 等）。"""
        result = []
        for child in cell_node.childNodes:
            if child.nodeType != OdfNode.ELEMENT_NODE:
                continue
            ln = _local(child)
            if ln == 'p':
                sa = _para_style_attr(child)
                inline = _inline_html(child)
                result.append(f'<p{sa}>{inline or "<br/>"}</p>')
            elif ln == 'h':
                level = child.getAttribute('outlinelevel') or '1'
                sa = _para_style_attr(child)
                inline = _inline_html(child)
                result.append(f'<h{level}{sa}>{inline or "<br/>"}</h{level}>')
        return ''.join(result) if result else '<p>&nbsp;</p>'

    def _process_table(node):
        """將 ODT table 轉換為 HTML table（使用 <tbody>，全部 <td>，保留段落樣式）。"""
        rows_html = []
        for child in node.childNodes:
            ln = _local(child)
            if ln == 'table-header-rows':
                for row in child.childNodes:
                    if _local(row) == 'table-row':
                        cells = []
                        for c in row.childNodes:
                            if _local(c) == 'table-cell':
                                cells.append(f'<td>{_cell_paragraphs(c)}</td>')
                        if cells:
                            rows_html.append(f'<tr>{"".join(cells)}</tr>')
            elif ln == 'table-row':
                cells = []
                for c in child.childNodes:
                    if _local(c) == 'table-cell':
                        cells.append(f'<td>{_cell_paragraphs(c)}</td>')
                if cells:
                    rows_html.append(f'<tr>{"".join(cells)}</tr>')
        if rows_html:
            parts.append(
                '<table style="width:100%">'
                '<tbody>'
                + ''.join(rows_html)
                + '</tbody>'
                + '</table>'
            )

    def _process(nodes):
        for node in nodes:
            if not hasattr(node, 'qname'):
                continue
            ln = _local(node)
            sa = _para_style_attr(node)
            if ln == 'h':
                level = node.getAttribute('outlinelevel') or '1'
                inline = _inline_html(node)
                if inline:
                    parts.append(f'<h{level}{sa}>{inline}</h{level}>')
            elif ln == 'p':
                inline = _inline_html(node)
                parts.append(f'<p{sa}>{inline}</p>' if inline else '<p><br/></p>')
            elif ln == 'list':
                items = []
                for item in node.childNodes:
                    if _local(item) == 'list-item':
                        for sub in item.childNodes:
                            sln = _local(sub)
                            if sln in ('p', 'h'):
                                t = _inline_html(sub)
                                if t:
                                    items.append(f'<li>{t}</li>')
                if items:
                    parts.append(f'<ul>{"".join(items)}</ul>')
            elif ln == 'table':
                _process_table(node)
            else:
                # 容器節點（text, section, frame 等）→ 繼續往下
                if hasattr(node, 'childNodes') and node.childNodes:
                    _process(node.childNodes)

    _process(doc.body.childNodes)
    return '\n'.join(parts) if parts else '<p></p>'


def _lo_convert_to_html(file_bytes, ext):
    """使用 LibreOffice headless 將 ODT/DOCX 轉換為 HTML。

    若 LibreOffice 未安裝，回傳 None（由呼叫端 fallback 至 Python 解析器）。
    轉換後呼叫 _lo_postprocess() 修正表格寬度、字體、頁尾、空白頁等問題。
    """
    if not shutil.which('soffice'):
        return None

    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, f'input{ext}')
        with open(in_path, 'wb') as fp:
            fp.write(file_bytes)

        proc = subprocess.run(
            ['soffice', '--headless', '--norestore', '--nofirststartwizard',
             '--nolockcheck',
             '--convert-to', 'html:HTML', '--outdir', tmpdir, in_path],
            capture_output=True, timeout=120,
            env={**os.environ, 'HOME': tmpdir},   # 每次獨立 config 目錄，避免鎖定衝突
        )
        if proc.returncode != 0:
            raise Exception(
                f'LibreOffice 轉換失敗：'
                f'{proc.stderr.decode("utf-8", errors="replace")[:400]}'
            )
        html_path = os.path.join(tmpdir, 'input.html')
        if not os.path.exists(html_path):
            raise Exception('LibreOffice：找不到輸出 HTML 檔案')

        with open(html_path, 'r', encoding='utf-8', errors='replace') as fp:
            raw_html = fp.read()

    style_m = re.search(r'<style[^>]*>(.*?)</style>', raw_html, re.DOTALL | re.IGNORECASE)
    margins = _extract_page_margins(style_m.group(1)) if style_m else None
    return _lo_postprocess(raw_html), margins


def _extract_page_margins(style_text):
    """從 @page CSS 規則提取頁面邊距並轉換為 px（96 dpi）。
    支援 cm、mm、in、pt、px 單位。回傳 {'top':N,'right':N,'bottom':N,'left':N} 或 None。
    """
    page_m = re.search(r'@page\s*\{([^}]+)\}', style_text, re.IGNORECASE)
    if not page_m:
        return None

    def _to_px(val):
        val = val.strip()
        try:
            if val.endswith('cm'):  return round(float(val[:-2]) * 37.795)
            if val.endswith('mm'):  return round(float(val[:-2]) * 3.7795)
            if val.endswith('in'):  return round(float(val[:-2]) * 96)
            if val.endswith('pt'):  return round(float(val[:-2]) * 96 / 72)
            if val.endswith('px'):  return round(float(val[:-2]))
        except Exception:
            pass
        return None

    props = {}
    for decl in page_m.group(1).split(';'):
        if ':' not in decl:
            continue
        prop, _, val = decl.partition(':')
        props[prop.strip().lower()] = val.strip()

    margins = {}
    if 'margin' in props:
        parts = props['margin'].split()
        if len(parts) == 1:
            v = _to_px(parts[0])
            if v:
                margins = {'top': v, 'right': v, 'bottom': v, 'left': v}
        elif len(parts) == 2:
            v, h = _to_px(parts[0]), _to_px(parts[1])
            if v and h:
                margins = {'top': v, 'right': h, 'bottom': v, 'left': h}
        elif len(parts) == 4:
            vals = [_to_px(p) for p in parts]
            if all(vals):
                margins = {'top': vals[0], 'right': vals[1],
                           'bottom': vals[2], 'left': vals[3]}
    for side in ('top', 'right', 'bottom', 'left'):
        key = f'margin-{side}'
        if key in props:
            v = _to_px(props[key])
            if v:
                margins[side] = v

    return margins if margins else None


def _lo_postprocess(raw_html):
    """後處理 LibreOffice HTML 輸出，修正已知的格式問題：

    1. <font> → <span>：保留 style 與 color，避免 sanitizer 濾掉後喪失格式
    2. 表格寬度：移除固定 px 寬度，改為 width:100%（避免超出頁面邊界）
    3. page-break-before: always 首段：移除（避免在頂部產生空白頁）
    4. <div title="footer"> 整塊移除（文件原生頁尾不進編輯器）
    5. <sdfield>：移除標籤但保留文字（頁碼欄位）
    6. 巢狀 <ol><ol>：展平為單層（LibreOffice 用多層巢狀表示縮排）
    7. class → CSS inline（針對 body/p/span 等 class 對應 style block）
    8. 回傳 <body> 內容字串
    """
    from io import StringIO

    # 取得 <style> 與 <body>
    style_m = re.search(r'<style[^>]*>(.*?)</style>', raw_html, re.DOTALL | re.IGNORECASE)
    body_m  = re.search(r'<body[^>]*>(.*?)</body>',  raw_html, re.DOTALL | re.IGNORECASE)
    if not body_m:
        return raw_html
    body_str = body_m.group(1).strip()

    # ── CSS class map ────────────────────────────────────────────────────────
    css_map = {}
    if style_m:
        for m in re.finditer(r'([a-zA-Z0-9]*)\.([\w-]+)\s*\{([^}]+)\}',
                             style_m.group(1)):
            tag  = m.group(1).lower() or '*'
            cls  = m.group(2)
            props = ' '.join(m.group(3).split()).strip('; ')
            key  = (tag, cls)
            css_map[key] = ((css_map.get(key) or '') + '; ' + props).strip('; ')

    try:
        parser  = etree.HTMLParser(recover=True)
        doc     = etree.parse(StringIO(f'<div id="_lo_wrap_">{body_str}</div>'), parser)
        root    = doc.getroot()
        wrap    = root.find('.//*[@id="_lo_wrap_"]')
        if wrap is None:
            wrap = root.find('.//body') or root

        # ── Pass 1：移除 <div title="footer"> 整塊 & <sdfield> ─────────────
        for el in list(wrap.iter()):
            if not isinstance(el.tag, str):
                continue
            tag = el.tag.lower()

            if tag == 'div' and (el.get('title') or '').lower() == 'footer':
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
                continue

            if tag == 'sdfield':
                # 保留文字，移除標籤本身
                parent = el.getparent()
                if parent is not None:
                    idx = list(parent).index(el)
                    text = (el.text or '') + (el.tail or '')
                    el.tail = None
                    parent.remove(el)
                    if text:
                        if idx > 0:
                            prev = parent[idx - 1]
                            prev.tail = (prev.tail or '') + text
                        else:
                            parent.text = (parent.text or '') + text

        # ── Pass 2：<font> → <span>，保留 face/size/color ──────────────────
        _font_size_map = {
            '1': '8pt', '2': '10pt', '3': '12pt', '4': '14pt',
            '5': '18pt', '6': '24pt', '7': '36pt',
        }
        for el in list(wrap.iter('font')):
            el.tag = 'span'
            face  = el.attrib.pop('face', None)
            size  = el.attrib.pop('size', None)
            color = el.attrib.pop('color', None)

            css_parts = []
            if face:
                css_parts.append(f'font-family: {face}')
            if size:
                pt = _font_size_map.get(size.strip())
                if pt:
                    css_parts.append(f'font-size: {pt}')
            if color and color not in ('auto', '#000000', '#ffffff'):
                css_parts.append(f'color: {color}')
            if css_parts:
                existing = (el.get('style') or '').strip('; ')
                new_css  = '; '.join(css_parts)
                el.set('style', (existing + '; ' + new_css).strip('; ')
                                if existing else new_css)

        # ── Pass 3：CSS class → inline style ────────────────────────────
        for el in wrap.iter():
            if not isinstance(el.tag, str):
                continue
            tag     = el.tag.lower()
            cls_val = el.get('class', '')
            if not cls_val:
                continue
            extras = []
            for cls in cls_val.split():
                p = css_map.get((tag, cls)) or css_map.get(('*', cls))
                if p:
                    extras.append(p)
            if extras:
                existing  = (el.get('style') or '').strip('; ')
                # reversed(extras)：class 屬性中「第一個 class」排在字串末尾，
                # 使其在重複屬性時優先（CSS last-wins 語意）。
                # existing 放最後：原始 inline style 最優先（模擬 inline > class 層級）。
                new_style = '; '.join(reversed(extras))
                el.set('style', (new_style + '; ' + existing).strip('; ')
                                if existing else new_style)
            del el.attrib['class']

        # ── Pass 4：移除首個元素的 page-break-before: always ────────────
        _first_block_fixed = [False]
        for el in wrap:
            if not isinstance(el.tag, str):
                continue
            if not _first_block_fixed[0]:
                style = el.get('style', '')
                new_style = re.sub(
                    r';\s*page-break-before\s*:\s*always|page-break-before\s*:\s*always\s*;?',
                    '', style, flags=re.IGNORECASE
                ).strip('; ')
                if new_style != style:
                    el.set('style', new_style) if new_style else el.attrib.pop('style', None)
                _first_block_fixed[0] = True
                break

        # ── Pass 5：表格寬度處理 ──────────────────────────────────────────
        # 表格本身：移除絕對 px 寬度，改為 width:100%（尊重使用者邊距）
        # 欄寬：讀取 <col width="Npx">，換算成百分比後保留相對比例
        for tbl in wrap.iter('table'):
            tbl_width_px = 0
            try:
                tbl_width_px = int(tbl.get('width') or 0)
            except (ValueError, TypeError):
                pass
            tbl.attrib.pop('width', None)
            tbl.attrib.pop('cellpadding', None)
            tbl.attrib.pop('cellspacing', None)
            existing = (tbl.get('style') or '').strip('; ')
            tbl.set('style', (existing + '; width: 100%; border-collapse: collapse').strip('; ')
                             if existing else 'width: 100%; border-collapse: collapse')

            # 讀取 <col> 欄寬並換算成百分比
            cols = tbl.findall('.//col')
            if cols and tbl_width_px > 0:
                col_widths = []
                for col in cols:
                    try:
                        w = int(col.get('width') or 0)
                    except (ValueError, TypeError):
                        w = 0
                    col_widths.append(w)
                total_w = sum(col_widths)
                if total_w > 0:
                    for col, w in zip(cols, col_widths):
                        col.attrib.pop('width', None)
                        pct = round(w / total_w * 100, 1)
                        col.set('style', f'width: {pct}%')
            else:
                # 沒有 <col> 或無總寬資訊：移除個別欄的固定寬度
                for col in cols:
                    col.attrib.pop('width', None)

        # ── Pass 6：展平巢狀 <ol>（只含單個子 <ol> 的層級直接跳過）────
        def _flatten_lists(parent):
            for child in list(parent):
                if not isinstance(child.tag, str):
                    continue
                # 若 <ol> 只含一個直接子 <ol>，用內層取代外層
                if child.tag.lower() in ('ol', 'ul'):
                    real_children = [c for c in child
                                     if isinstance(c.tag, str) and c.tag.lower()
                                     not in ('', None.__class__.__name__)]
                    if (len(real_children) == 1 and
                            real_children[0].tag.lower() in ('ol', 'ul')):
                        idx = list(parent).index(child)
                        inner = real_children[0]
                        child.remove(inner)
                        parent.remove(child)
                        parent.insert(idx, inner)
                        _flatten_lists(parent)
                        return
                    _flatten_lists(child)
                else:
                    _flatten_lists(child)
        _flatten_lists(wrap)

        # ── 序列化 ────────────────────────────────────────────────────────
        out = etree.tostring(wrap, encoding='unicode', method='html')
        # 去掉外層 wrapper div
        out = re.sub(r'^<div[^>]*>', '', out)
        out = re.sub(r'</div>$', '', out)
        return out.strip()

    except Exception:
        return body_str


class DocEditorController(http.Controller):

    @http.route('/dobtor_doc/load', type='json', auth='user', methods=['POST'])
    def load_document(self, doc_id, **kw):
        """載入文件資料（HTML 內容＋頁面設定）。"""
        doc = request.env['doc.document'].browse(doc_id)
        doc.check_access_rule('read')
        return {
            'id': doc.id,
            'name': doc.name,
            'content_html': doc.get_content_html(),
            'header_html': doc.header_html or '',
            'footer_html': doc.footer_html or '',
            'page_format': doc.page_format,
            'margin_top': doc.margin_top,
            'margin_bottom': doc.margin_bottom,
            'margin_left': doc.margin_left,
            'margin_right': doc.margin_right,
            'model_id': doc.model_id.id if doc.model_id else False,
            'model_name': doc.model_id.model if doc.model_id else False,
            'has_different_first_page': doc.has_different_first_page,
            'first_header_html': doc.first_header_html or '',
            'first_footer_html': doc.first_footer_html or '',
        }

    @http.route('/dobtor_doc/save', type='json', auth='user', methods=['POST'])
    def save_document(self, doc_id=None, content_html=None, header_html=None,
                      footer_html=None, name=None, **kw):
        """儲存文件 HTML 內容。"""
        if not doc_id:
            return {'success': False, 'error': 'doc_id required'}
        doc = request.env['doc.document'].browse(doc_id)
        doc.check_access_rule('write')
        vals = {}
        if content_html is not None:
            vals['content_html'] = content_html
        if header_html is not None:
            vals['header_html'] = header_html
        if footer_html is not None:
            vals['footer_html'] = footer_html
        if name is not None:
            vals['name'] = name
        if vals:
            doc.write(vals)
        return {'success': True, 'write_date': doc.write_date.isoformat()}

    @http.route('/dobtor_doc/save_settings', type='json', auth='user', methods=['POST'])
    def save_settings(self, doc_id, **kw):
        """儲存頁面格式與邊距設定。"""
        doc = request.env['doc.document'].browse(doc_id)
        doc.check_access_rule('write')
        allowed = ('page_format', 'margin_top', 'margin_bottom',
                   'margin_left', 'margin_right',
                   'default_column_count', 'default_column_gap', 'column_rule_style')
        vals = {k: v for k, v in kw.items() if k in allowed}
        if vals:
            doc.write(vals)
        return {'success': True}

    @http.route('/dobtor_doc/fields', type='json', auth='user', methods=['POST'])
    def get_fields(self, model_name, doc_id=None, **kw):
        """取得指定模型的可用欄位清單。"""
        mixin = request.env['doc.render.mixin']
        try:
            return mixin.get_available_fields(model_name)
        except Exception as e:
            return {'error': str(e)}

    @http.route('/dobtor_doc/render_preview', type='json', auth='user', methods=['POST'])
    def render_preview(self, doc_id, record_model, record_id, **kw):
        """將欄位變數渲染為實際值（預覽用）。"""
        doc = request.env['doc.document'].browse(doc_id)
        doc.check_access_rule('read')
        try:
            record = request.env[record_model].browse(record_id)
            record.check_access_rule('read')
            rendered = doc._render_template(doc.get_content_html(), record)
            return {'html': rendered}
        except Exception as e:
            return {'error': str(e)}

    @http.route('/dobtor_doc/export', type='json', auth='user', methods=['POST'])
    def export_document(self, doc_id, format='pdf', quality='high',
                        record_model=None, record_id=None, **kw):
        """匯出文件為 PDF 或 DOCX。quality: 'high'（後端）"""
        doc = request.env['doc.document'].browse(doc_id)
        doc.check_access_rule('read')

        record = None
        if record_model and record_id:
            try:
                record = request.env[record_model].browse(record_id)
                record.check_access_rule('read')
            except Exception:
                record = None

        try:
            if format == 'pdf':
                file_bytes = doc._generate_pdf(record=record)
                mimetype = 'application/pdf'
                filename = f'{doc.name}.pdf'
            elif format == 'docx':
                file_bytes = doc._generate_docx_via_libreoffice(record=record)
                mimetype = ('application/vnd.openxmlformats-officedocument'
                            '.wordprocessingml.document')
                filename = f'{doc.name}.docx'
            else:
                return {'error': f'不支援的格式：{format}'}

            return {
                'filename': filename,
                'data': base64.b64encode(file_bytes).decode(),
                'mimetype': mimetype,
            }
        except Exception as e:
            return {'error': str(e)}

    @http.route('/dobtor_doc/save_version', type='json', auth='user', methods=['POST'])
    def save_version(self, doc_id, label=None, **kw):
        """儲存版本快照。"""
        doc = request.env['doc.document'].browse(doc_id)
        doc.check_access_rule('write')
        doc.action_save_version(label=label)
        return {'success': True}

    @http.route('/dobtor_doc/import', type='http', auth='user', methods=['POST'], csrf=False)
    def import_document(self, **kw):
        """匯入 DOCX / ODT 檔案，轉換為 HTML。
        DOCX：使用 mammoth（純 Python）
        ODT：使用 odfpy（純 Python）
        其他格式：嘗試 LibreOffice，不可用時回傳錯誤
        """
        def _json_resp(data):
            return request.make_response(
                json.dumps(data, ensure_ascii=False),
                headers={'Content-Type': 'application/json; charset=utf-8'},
            )

        upload = request.httprequest.files.get('file')
        if not upload:
            return _json_resp({'error': '未收到檔案'})

        filename = upload.filename or ''
        ext = os.path.splitext(filename)[1].lower()

        try:
            file_bytes = upload.read()
            page_margins = None

            if ext in ('.docx', '.odt'):
                # ── 優先使用 LibreOffice（格式保留最完整）──
                lo_result = _lo_convert_to_html(file_bytes, ext)
                if lo_result is not None:
                    body_html, page_margins = lo_result
                elif ext == '.docx':
                    # Fallback：zipfile + lxml（LibreOffice 未安裝時）
                    body_html = _docx_to_html_with_format(file_bytes)
                else:
                    # Fallback：odfpy（LibreOffice 未安裝時）
                    body_html = _odt_to_html(file_bytes)

            else:
                # ── 其他格式：LibreOffice（僅支援已安裝時）──
                if not shutil.which('soffice'):
                    return _json_resp({
                        'error': f'不支援的格式（{ext}）。支援格式：.docx、.odt'
                    })
                lo_result = _lo_convert_to_html(file_bytes, ext)
                if lo_result is None:
                    return _json_resp({'error': f'LibreOffice 無法轉換格式：{ext}'})
                body_html, page_margins = lo_result

            resp = {'html': body_html}
            if page_margins:
                resp['margins'] = page_margins
            return _json_resp(resp)

        except Exception as e:
            return _json_resp({'error': str(e)})
