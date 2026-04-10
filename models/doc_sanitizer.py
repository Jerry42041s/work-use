import re
from odoo import models

try:
    from lxml import etree
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

# 允許的 HTML 標籤白名單
ALLOWED_TAGS = {
    'p', 'div', 'span', 'br', 'hr',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'blockquote', 'pre', 'code',
    'ul', 'ol', 'li',
    'table', 'colgroup', 'col', 'thead', 'tbody', 'tfoot', 'tr', 'th', 'td', 'caption',
    'strong', 'b', 'em', 'i', 'u', 's', 'del', 'ins', 'sub', 'sup', 'mark',
    'a', 'img',
}

# 允許的屬性（按標籤）
ALLOWED_ATTRIBUTES = {
    '*': ['class', 'style', 'id', 'title', 'dir', 'lang'],
    'a': ['href', 'target', 'rel'],
    'img': ['src', 'alt', 'width', 'height'],
    'td': ['colspan', 'rowspan'],
    'th': ['colspan', 'rowspan', 'scope'],
    'col': ['span'],
    'colgroup': ['span'],
    # 清單 HTML 屬性（LibreOffice 及自訂清單類型）
    'ol': ['type', 'start'],
    'ul': ['type'],
    'li': ['value'],
    # 文件編輯器自訂屬性
    'span': ['data-expression', 'data-field', 'data-field-label', 'contenteditable'],
    'div': ['data-expression', 'data-field', 'data-field-label', 'contenteditable',
            'data-page-break', 'data-page', 'data-page-gap'],
}

# 允許的 CSS 屬性
ALLOWED_CSS = {
    'color', 'background-color', 'font-size', 'font-family', 'font-weight',
    'font-style', 'text-align', 'text-decoration', 'line-height',
    # 縮排與間距（LibreOffice 大量使用）
    'text-indent', 'letter-spacing', 'word-spacing', 'text-transform',
    'margin', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right',
    'padding', 'padding-top', 'padding-bottom', 'padding-left', 'padding-right',
    'border', 'border-top', 'border-bottom', 'border-left', 'border-right',
    'border-collapse', 'border-spacing', 'border-color', 'border-style', 'border-width',
    'width', 'height', 'min-width', 'max-width', 'min-height',
    'display', 'vertical-align', 'white-space',
    # 分頁控制
    'page-break-before', 'page-break-after', 'page-break-inside',
    'break-before', 'break-after', 'break-inside',
    # 清單樣式
    'list-style-type', 'list-style', 'list-style-position', 'list-style-image',
    # 多欄排版
    'column-count', 'column-gap',
    # 圖文混排
    'float', 'clear',
    # 版面
    'box-sizing',
}

# 危險 CSS 模式：阻擋 expression/javascript/vbscript，以及非 data:image/ 的 url()
CSS_DANGEROUS = re.compile(r'(expression|javascript|vbscript)\s*\(', re.IGNORECASE)
# url() 需單獨判斷，允許 data:image/ 通過（用於 base64 圖片與 list-style-image）
_CSS_URL = re.compile(r'url\s*\(', re.IGNORECASE)
_CSS_SAFE_DATA_URL = re.compile(r'url\s*\(\s*["\']?\s*data:image/', re.IGNORECASE)


class DocSanitizer(models.AbstractModel):
    _name = 'doc.sanitizer'
    _description = '文件 HTML 清理器'

    def sanitize_html(self, html_content):
        """清理 HTML，保留白名單標籤與自訂 data-* 屬性，移除危險內容。"""
        if not html_content:
            return html_content

        if not HAS_LXML:
            # lxml 不可用時，用基本 regex 清理最危險的內容
            html_content = re.sub(r'<script[^>]*>.*?</script>', '', html_content,
                                  flags=re.DOTALL | re.IGNORECASE)
            html_content = re.sub(r'\son\w+\s*=', ' data-removed=', html_content,
                                  flags=re.IGNORECASE)
            return html_content

        try:
            from io import StringIO
            parser = etree.HTMLParser(recover=True)
            doc = etree.parse(StringIO(f'<div>{html_content}</div>'), parser)
            root = doc.getroot()
            self._walk_and_sanitize(root)
            body = root.find('.//body/div')
            if body is not None:
                result = etree.tostring(body, encoding='unicode', method='html')
                # 移除外層 <div> 包裝
                result = re.sub(r'^<div[^>]*>', '', result)
                result = re.sub(r'</div>$', '', result)
                return result
        except Exception:
            pass
        return html_content

    def _walk_and_sanitize(self, element):
        to_remove = []
        for child in list(element):
            tag = child.tag.lower() if isinstance(child.tag, str) else ''

            # 危險標籤直接移除
            if tag in ('script', 'iframe', 'object', 'embed', 'applet',
                       'form', 'input', 'textarea', 'button', 'select'):
                to_remove.append(child)
                continue

            if tag not in ALLOWED_TAGS:
                # 非白名單標籤：保留文字內容（unwrap）
                parent = child.getparent()
                if parent is not None:
                    idx = list(parent).index(child)
                    if child.text:
                        prev = parent[idx - 1] if idx > 0 else None
                        if prev is not None:
                            prev.tail = (prev.tail or '') + child.text
                        else:
                            parent.text = (parent.text or '') + child.text
                    for subchild in reversed(list(child)):
                        parent.insert(idx, subchild)
                    to_remove.append(child)
                    continue

            self._sanitize_attributes(child, tag)
            self._walk_and_sanitize(child)

        for child in to_remove:
            if child.getparent() is not None:
                child.getparent().remove(child)

    def _sanitize_attributes(self, element, tag):
        allowed = set(ALLOWED_ATTRIBUTES.get('*', []))
        allowed.update(ALLOWED_ATTRIBUTES.get(tag, []))

        to_remove = []
        for attr in list(element.attrib):
            attr_lower = attr.lower()
            # 移除所有 on* 事件處理器
            if attr_lower.startswith('on'):
                to_remove.append(attr)
                continue
            # 非白名單屬性移除（但保留 data-* 用於自訂元素）
            if attr_lower not in allowed and not attr_lower.startswith('data-'):
                to_remove.append(attr)
                continue
            # 檢查危險 URL
            if attr_lower in ('href', 'src', 'action'):
                val = element.attrib[attr]
                if re.match(r'\s*(javascript|vbscript|data\s*:(?!image/))',
                            val, re.IGNORECASE):
                    to_remove.append(attr)
                    continue
            # 清理 style
            if attr_lower == 'style':
                element.attrib[attr] = self._sanitize_style(element.attrib[attr])

        for attr in to_remove:
            del element.attrib[attr]

    def _sanitize_style(self, style_value):
        if CSS_DANGEROUS.search(style_value):
            return ''
        safe = []
        for decl in style_value.split(';'):
            decl = decl.strip()
            if ':' not in decl:
                continue
            prop, _, val = decl.partition(':')
            prop, val = prop.strip().lower(), val.strip()
            if prop not in ALLOWED_CSS:
                continue
            # 阻擋一般危險模式
            if CSS_DANGEROUS.search(val):
                continue
            # url() 特殊判斷：允許 data:image/，阻擋其他（包含 javascript:、外部 URL 等）
            if _CSS_URL.search(val) and not _CSS_SAFE_DATA_URL.search(val):
                continue
            safe.append(f'{prop}: {val}')
        return '; '.join(safe)
