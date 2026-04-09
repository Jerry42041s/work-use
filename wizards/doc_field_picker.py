from odoo import models, fields, api


class DocFieldPickerWizard(models.TransientModel):
    _name = 'doc.field.picker.wizard'
    _description = '欄位選擇器'

    document_id = fields.Many2one('doc.document', string='文件', required=True)
    model_id = fields.Many2one(
        'ir.model',
        related='document_id.model_id',
        readonly=True,
        string='目標模型',
    )
    selected_field_id = fields.Many2one(
        'ir.model.fields',
        domain="[('model_id', '=', model_id), ('store', '=', True),"
               " ('ttype', 'not in', ['binary', 'html', 'serialized'])]",
        string='選擇欄位',
    )
    expression_preview = fields.Char(
        string='運算式預覽',
        compute='_compute_preview',
    )

    @api.depends('selected_field_id')
    def _compute_preview(self):
        for rec in self:
            if rec.selected_field_id:
                rec.expression_preview = '{{ object.%s }}' % rec.selected_field_id.name
            else:
                rec.expression_preview = ''

    def action_copy_and_close(self):
        """返回運算式給前端複製使用。"""
        return {
            'type': 'ir.actions.act_window_close',
        }
