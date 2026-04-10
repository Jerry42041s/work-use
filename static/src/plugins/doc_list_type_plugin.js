/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { withSequence } from "@html_editor/utils/resource";
import { toolbarButtonProps } from "@html_editor/main/toolbar/toolbar";
import { _t } from "@web/core/l10n/translation";
import { Component, useState, reactive } from "@odoo/owl";
import { Dropdown } from "@web/core/dropdown/dropdown";
import { DropdownItem } from "@web/core/dropdown/dropdown_item";

// 清單類型選項
const LIST_TYPES = [
    { name: "• 圓點",       value: "disc",        tag: "ul", htmlType: null },
    { name: "○ 空心圓",    value: "circle",      tag: "ul", htmlType: null },
    { name: "■ 方塊",       value: "square",      tag: "ul", htmlType: null },
    { name: "1. 數字",      value: "decimal",     tag: "ol", htmlType: "1" },
    { name: "a. 小寫英文", value: "lower-alpha",  tag: "ol", htmlType: "a" },
    { name: "A. 大寫英文", value: "upper-alpha",  tag: "ol", htmlType: "A" },
    { name: "i. 小寫羅馬", value: "lower-roman",  tag: "ol", htmlType: "i" },
    { name: "I. 大寫羅馬", value: "upper-roman",  tag: "ol", htmlType: "I" },
];

class ListTypeSelector extends Component {
    static template = "dobtor_doc_editor.ListTypeSelector";
    static components = { Dropdown, DropdownItem };
    static props = {
        getItems: Function,
        getDisplay: Function,
        onSelected: Function,
        ...toolbarButtonProps,
    };

    setup() {
        this.items = this.props.getItems();
        // getDisplay() 回傳 reactive 物件，useState 使其可被觀察
        this.state = useState(this.props.getDisplay());
    }

    onSelected(item) {
        this.props.onSelected(item);
    }
}

export class DocListTypePlugin extends Plugin {
    static id = "docListType";
    static dependencies = ["selection"];

    setup() {
        this.listTypeState = reactive({ displayName: "清單" });
    }

    get resources() {
        return {
            toolbar_groups: [withSequence(65, { id: "docListType" })],
            toolbar_items: [
                {
                    id: "listType",
                    groupId: "docListType",
                    title: _t("清單類型"),
                    Component: ListTypeSelector,
                    props: {
                        getItems: () => LIST_TYPES,
                        getDisplay: () => this.listTypeState,
                        onSelected: (item) => this._applyListType(item),
                    },
                },
            ],
            // 游標移動時更新 Toolbar 顯示
            selectionchange_handlers: [this._updateDisplay.bind(this)],
        };
    }

    _updateDisplay() {
        const sel = this.dependencies.selection.getEditableSelection?.();
        if (!sel?.anchorNode) {
            this.listTypeState.displayName = "清單";
            return;
        }
        const el = sel.anchorNode.nodeType === 3
            ? sel.anchorNode.parentElement
            : sel.anchorNode;
        const listEl = el?.closest?.("ul, ol");
        if (listEl) {
            const lsType = listEl.style.listStyleType;
            const match = LIST_TYPES.find((t) => t.value === lsType);
            this.listTypeState.displayName = match ? match.name : "清單";
        } else {
            this.listTypeState.displayName = "清單";
        }
    }

    _applyListType(item) {
        const sel = this.dependencies.selection.getEditableSelection?.();
        if (!sel) return;

        const el = sel.anchorNode?.nodeType === 3
            ? sel.anchorNode.parentElement
            : sel.anchorNode;
        if (!el) return;

        const listEl = el.closest?.("ul, ol");
        if (!listEl) return;

        // 設定 list-style-type（DOM mutation 自動被 MutationObserver 記錄）
        listEl.style.listStyleType = item.value;
        if (item.htmlType) {
            listEl.setAttribute("type", item.htmlType);
        } else {
            listEl.removeAttribute("type");
        }

        // 若 ul/ol 標籤需要轉換（如 ul → ol）
        if (listEl.tagName.toLowerCase() !== item.tag) {
            this._convertListTag(listEl, item.tag);
        }
        this._updateDisplay();
    }

    /**
     * 將 <ul> 轉為 <ol>（或反之），保留所有屬性與子節點。
     */
    _convertListTag(listEl, newTag) {
        const newList = this.document.createElement(newTag);
        for (const attr of listEl.attributes) {
            newList.setAttribute(attr.name, attr.value);
        }
        while (listEl.firstChild) {
            newList.appendChild(listEl.firstChild);
        }
        listEl.parentNode.replaceChild(newList, listEl);
        return newList;
    }
}
