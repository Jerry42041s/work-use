/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { _t } from "@web/core/l10n/translation";
import { withSequence } from "@html_editor/utils/resource";
import { DocFieldPickerDialog } from "../../components/doc_field_picker/doc_field_picker";

/**
 * DocOdooFieldPlugin — 提供在文件中插入 Odoo 欄位變數的功能。
 * 透過 OWL Dialog（DocFieldPickerDialog）顯示欄位清單，
 * 使用者選取後插入 {{ object.field }} chip span。
 */
export class DocOdooFieldPlugin extends Plugin {
    static id = "docOdooField";
    static dependencies = ["history", "selection"];

    // Odoo 18 html_editor Plugin 透過 this.getService() 存取 service
    get _dialogService() {
        return this.getService?.("dialog") ?? this.services?.dialog ?? null;
    }

    get resources() {
        return {
            powerbox_categories: [
                withSequence(50, {
                    id: "docFields",
                    name: _t("欄位變數"),
                }),
            ],
            user_commands: [
                {
                    id: "insertFieldDialog",
                    title: _t("插入欄位變數"),
                    description: _t("從清單中選取 Odoo 欄位插入"),
                    icon: "fa-database",
                    run: () => this._openFieldPickerDialog(),
                },
                {
                    id: "insertCustomField",
                    title: _t("輸入自訂欄位"),
                    description: _t("手動輸入 object.field 路徑"),
                    icon: "fa-code",
                    run: () => this._promptAndInsertField(),
                },
            ],
            powerbox_items: [
                {
                    categoryId: "docFields",
                    commandId: "insertFieldDialog",
                    title: _t("插入欄位變數"),
                    description: _t("從欄位清單選取"),
                    icon: "fa-database",
                    keywords: ["field", "欄位", "變數", "odoo"],
                },
                {
                    categoryId: "docFields",
                    commandId: "insertCustomField",
                    title: _t("自訂欄位..."),
                    description: _t("手動輸入欄位路徑"),
                    icon: "fa-pencil",
                    keywords: ["custom", "自訂"],
                },
            ],
        };
    }

    _openFieldPickerDialog() {
        const dialogService = this._dialogService;
        if (!dialogService) {
            // fallback：dialog service 不可用時改用 prompt
            this._promptAndInsertField();
            return;
        }
        dialogService.add(DocFieldPickerDialog, {
            modelName: this.config?.dobtor_model_name || "",
            docId: this.config?.dobtor_doc_id || null,
            onInsert: (expression, label) => {
                this.insertFieldToken(expression, label);
            },
        });
    }

    _promptAndInsertField() {
        const fieldName = window.prompt(
            _t("請輸入欄位名稱（例如：name、date、partner_id.name）：")
        );
        if (fieldName && fieldName.trim()) {
            const expr = `{{ object.${fieldName.trim()} }}`;
            this.insertFieldToken(expr, fieldName.trim());
        }
    }

    /**
     * 在游標位置插入欄位 chip span。
     */
    insertFieldToken(expression, label) {
        const token = this.document.createElement("span");
        token.className = "doc-field-token";
        token.setAttribute("contenteditable", "false");
        token.setAttribute("data-expression", expression);
        token.setAttribute("data-field-label", label);
        token.textContent = expression;

        const selection = this.document.getSelection();
        if (selection && selection.rangeCount > 0) {
            const range = selection.getRangeAt(0);
            range.deleteContents();
            range.insertNode(token);

            const newRange = this.document.createRange();
            newRange.setStartAfter(token);
            newRange.collapse(true);
            selection.removeAllRanges();
            selection.addRange(newRange);
        } else if (this.editable) {
            this.editable.appendChild(token);
        }

        if (this.shared.addStep) this.shared.addStep();
    }
}
