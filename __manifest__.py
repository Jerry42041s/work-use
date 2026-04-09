{
    'name': 'Dobtor Doc Editor',
    'version': '18.0.2.1.0',
    'summary': 'Google Docs 等級的 native Odoo 文件編輯器',
    'description': """
Dobtor Doc Editor v2.1
======================
基於 Odoo 18 html_editor 建立的全功能文件編輯器：
- A4/A3/A5/Letter/Legal 頁面排版
- 自動分頁引擎（PaginationEngine）+ 手動分頁
- 頁首/頁尾（Singleton 模式）
- 欄位變數插入（OWL Dialog 選擇器）
- 多欄排版（2/3 欄 Section）
- AutoSave（Debounce + MaxWait + Idle）
- 離線緩存（OfflineManager）
- PDF / DOCX 匯出（後端 LibreOffice 高品質 + 前端草稿）
- DOCX / ODT 匯入
- 版本快照
- 多公司隔離
    """,
    'category': 'Productivity',
    'author': 'Dobtor',
    'license': 'LGPL-3',
    'depends': [
        'base',
        'web',
        'mail',
        'html_editor',
        'bus',
    ],
    'data': [
        'security/doc_groups.xml',
        'security/ir.model.access.csv',
        'security/doc_security.xml',
        'wizards/doc_field_picker_views.xml',
        'views/doc_document_views.xml',
        'views/doc_template_views.xml',
        'views/menu.xml',
        'data/doc_template_data.xml',
    ],
    'assets': {
        'web.assets_backend': [
            # CSS
            'dobtor_doc_editor/static/src/css/doc_editor.css',
            # Core 模組（PaginationEngine, AutoSave, etc.）
            'dobtor_doc_editor/static/src/core/pagination_engine.js',
            'dobtor_doc_editor/static/src/core/auto_save_manager.js',
            'dobtor_doc_editor/static/src/core/leader_election.js',
            'dobtor_doc_editor/static/src/core/offline_manager.js',
            'dobtor_doc_editor/static/src/core/lazy_loader.js',
            # Plugins
            'dobtor_doc_editor/static/src/js/plugins/doc_page_format_plugin.js',
            'dobtor_doc_editor/static/src/js/plugins/doc_export_plugin.js',
            'dobtor_doc_editor/static/src/plugins/doc_multi_column_plugin.js',
            'dobtor_doc_editor/static/src/plugins/doc_font_family_plugin.js',
            'dobtor_doc_editor/static/src/plugins/doc_line_height_plugin.js',
            'dobtor_doc_editor/static/src/plugins/doc_table_merge_plugin.js',
            'dobtor_doc_editor/static/src/plugins/doc_formatting_plugins.xml',
            # Components（依賴順序）
            'dobtor_doc_editor/static/src/components/doc_field_picker/doc_field_picker.xml',
            'dobtor_doc_editor/static/src/components/doc_field_picker/doc_field_picker.js',
            'dobtor_doc_editor/static/src/js/plugins/doc_odoo_field_plugin.js',
            'dobtor_doc_editor/static/src/components/doc_ruler/doc_ruler.xml',
            'dobtor_doc_editor/static/src/components/doc_ruler/doc_ruler.js',
            'dobtor_doc_editor/static/src/components/doc_page_layout/doc_page_layout.xml',
            'dobtor_doc_editor/static/src/components/doc_page_layout/doc_page_layout.js',
            'dobtor_doc_editor/static/src/components/doc_editor/doc_editor.xml',
            'dobtor_doc_editor/static/src/components/doc_editor/doc_editor.js',
        ],
    },
    'installable': True,
    'application': True,
    'auto_install': False,
}
