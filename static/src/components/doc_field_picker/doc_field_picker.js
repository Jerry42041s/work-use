/** @odoo-module **/

import { Component, useState, onWillStart } from "@odoo/owl";
import { Dialog } from "@web/core/dialog/dialog";
import { rpc } from "@web/core/network/rpc";
import { _t } from "@web/core/l10n/translation";

/**
 * DocFieldPickerDialog — OWL Dialog 欄位選擇器。
 * 從 Powerbox 觸發，顯示目前文件可用的 Odoo 欄位，選取後插入游標位置。
 */
export class DocFieldPickerDialog extends Component {
    static template = "dobtor_doc_editor.DocFieldPickerDialog";
    static components = { Dialog };
    static props = {
        modelName: { type: String, optional: true },
        docId: { type: Number, optional: true },
        onInsert: { type: Function },
        close: { type: Function },
    };

    setup() {
        this.state = useState({
            fields: [],
            loading: true,
            error: null,
            search: "",
            expanded: new Set(),
        });

        onWillStart(async () => {
            await this._loadFields();
        });
    }

    async _loadFields() {
        if (!this.props.modelName) {
            this.state.loading = false;
            this.state.error = _t("此文件未設定目標 Model，無法列出欄位。");
            return;
        }
        try {
            const fields = await rpc("/dobtor_doc/fields", {
                model_name: this.props.modelName,
                doc_id: this.props.docId || null,
            });
            this.state.fields = fields;
        } catch (e) {
            this.state.error = e.message || _t("載入欄位失敗");
        } finally {
            this.state.loading = false;
        }
    }

    get filteredFields() {
        const q = this.state.search.trim().toLowerCase();
        if (!q) return this.state.fields;
        return this.state.fields.filter(f =>
            f.name.toLowerCase().includes(q) ||
            (f.label || "").toLowerCase().includes(q)
        );
    }

    onSearchInput(ev) {
        this.state.search = ev.target.value;
    }

    toggleExpand(fieldName) {
        if (this.state.expanded.has(fieldName)) {
            this.state.expanded.delete(fieldName);
        } else {
            this.state.expanded.add(fieldName);
        }
        // 觸發重新渲染
        this.state.expanded = new Set(this.state.expanded);
    }

    insertField(expression, label) {
        this.props.onInsert(expression, label);
        this.props.close();
    }

    insertSubField(parentName, field) {
        const expr = `{{ object.${parentName}.${field.name} }}`;
        const label = `${parentName}.${field.name}`;
        this.insertField(expr, label);
    }
}
