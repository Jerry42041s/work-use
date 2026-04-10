from odoo import models


class IrActionsReport(models.Model):
    _inherit = 'ir.actions.report'

    def _build_wkhtmltopdf_args(
        self, paperformat_id, landscape,
        specific_paperformat_args=None, set_viewport_size=False
    ):
        """擴充 specific_paperformat_args 支援 data-report-margin-left/right。

        Odoo 18 原始實作（ir_actions_report.py 第 344、351 行）將 --margin-left/right
        硬寫為 paperformat_id.margin_left/right，忽略 specific_paperformat_args 中的
        data-report-margin-left/right 鍵值。
        本覆寫在 super() 執行後，將 args 清單中對應的值替換為指定值。
        """
        args = super()._build_wkhtmltopdf_args(
            paperformat_id, landscape, specific_paperformat_args, set_viewport_size
        )
        if not specific_paperformat_args:
            return args

        for side in ('left', 'right'):
            key = f'data-report-margin-{side}'
            cli_flag = f'--margin-{side}'
            if key in specific_paperformat_args and cli_flag in args:
                idx = args.index(cli_flag)
                args[idx + 1] = str(specific_paperformat_args[key])
        return args
