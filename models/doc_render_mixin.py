from odoo import models, api
from odoo.exceptions import UserError, AccessError


class DocRenderMixin(models.AbstractModel):
    _name = 'doc.render.mixin'
    _description = '文件渲染 Mixin'

    def _render_template(self, html, record):
        """使用 Jinja2 SandboxedEnvironment 渲染 {{ field }} 變數，防止 SSTI。"""
        if not html or not record:
            return html or ''
        try:
            from jinja2.sandbox import SandboxedEnvironment
            env = SandboxedEnvironment()
            template = env.from_string(html)
            return template.render(object=record, user=self.env.user)
        except Exception as e:
            return html  # 渲染失敗時回傳原始 HTML

    @api.model
    def get_available_fields(self, model_name, max_depth=2):
        """回傳指定 model 的可用欄位清單（含 Many2one 子欄位，深度限制 2）。"""
        if model_name not in self.env:
            raise UserError(f"模型 '{model_name}' 不存在")

        try:
            self.env[model_name].check_access_rights('read')
        except AccessError:
            raise AccessError(f"您沒有讀取 '{model_name}' 的權限")

        return self._get_fields_for_model(model_name, max_depth=max_depth, current_depth=0)

    def _get_fields_for_model(self, model_name, max_depth=2, current_depth=0):
        """遞迴取得模型欄位，限制深度以防止無限遞迴。"""
        IrModelFields = self.env['ir.model.fields']
        fields = IrModelFields.search([
            ('model', '=', model_name),
            ('store', '=', True),
            ('ttype', 'in', [
                'char', 'text', 'integer', 'float', 'monetary',
                'date', 'datetime', 'boolean', 'selection', 'many2one',
            ]),
        ], order='field_description asc')

        result = []
        for f in fields:
            field_info = {
                'name': f.name,
                'label': f.field_description,
                'type': f.ttype,
                'expression': '{{{{ object.{} }}}}'.format(f.name),
            }
            # Many2one：若未達深度限制，遞迴取子欄位
            if f.ttype == 'many2one' and f.relation and current_depth < max_depth:
                if f.relation in self.env:
                    try:
                        self.env[f.relation].check_access_rights('read')
                        field_info['relation'] = f.relation
                        field_info['sub_fields'] = self._get_fields_for_model(
                            f.relation, max_depth, current_depth + 1
                        )
                    except AccessError:
                        field_info['sub_fields'] = []
            result.append(field_info)
        return result
