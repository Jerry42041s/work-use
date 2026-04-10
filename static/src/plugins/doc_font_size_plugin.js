/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { withSequence } from "@html_editor/utils/resource";
import { toolbarButtonProps } from "@html_editor/main/toolbar/toolbar";
import { _t } from "@web/core/l10n/translation";
import { Component, useState, reactive } from "@odoo/owl";
import { Dropdown } from "@web/core/dropdown/dropdown";
import { DropdownItem } from "@web/core/dropdown/dropdown_item";

// 常用字體大小（pt）
const FONT_SIZES_PT = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];

/**
 * 解析 fontSize 字串（如 "11pt"、"14px"），回傳 { value, unit }。
 */
function parseFontSize(sizeStr) {
    if (!sizeStr) return null;
    const m = sizeStr.match(/^([\d.]+)\s*(pt|px|em|rem)$/i);
    if (!m) return null;
    return { value: parseFloat(m[1]), unit: m[2].toLowerCase() };
}

/**
 * FontSizeInput — 字體大小輸入框 + Dropdown OWL Component
 */
class FontSizeInput extends Component {
    static template = "dobtor_doc_editor.FontSizeInput";
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

    onInputChange(event) {
        const raw = event.target.value.trim();
        if (raw) this.props.onSelected(raw);
    }

    onInputKeydown(event) {
        if (event.key === "Enter") {
            event.preventDefault();
            this.onInputChange(event);
        }
    }

    onDropdownSelected(ptValue) {
        this.props.onSelected(`${ptValue}pt`);
    }
}

/**
 * DocFontSizePlugin — 字體大小工具列插件
 *
 * 優先讀取 inline style 原始值（保留 pt 單位），
 * fallback 至 computed style（px，Odoo 預設行為）。
 */
export class DocFontSizePlugin extends Plugin {
    static id = "docFontSize";
    static dependencies = ["selection"];

    setup() {
        this.fontSizeState = reactive({ displayValue: "", unit: "pt" });
    }

    get resources() {
        return {
            toolbar_groups: [withSequence(25, { id: "docFontSize" })],
            toolbar_items: [
                {
                    id: "fontSize",
                    groupId: "docFontSize",
                    title: _t("字體大小"),
                    Component: FontSizeInput,
                    props: {
                        getItems: () => FONT_SIZES_PT,
                        getDisplay: () => this.fontSizeState,
                        onSelected: (sizeStr) => this._setFontSize(sizeStr),
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
            this.fontSizeState.displayValue = "";
            return;
        }
        const size = this._getCurrentFontSize(sel.anchorNode);
        this.fontSizeState.displayValue = size?.raw || "";
        this.fontSizeState.unit = size?.unit || "pt";
    }

    /**
     * 讀取節點的字體大小。
     * 優先回傳 inline style（保留 pt）；fallback 至 computed style（px）。
     */
    _getCurrentFontSize(anchorNode) {
        let el = anchorNode;
        if (el && el.nodeType === 3) el = el.parentElement;
        if (!el) return null;

        // 優先：inline style（保留 pt 原始值）
        const inlineSize = el.style?.fontSize;
        if (inlineSize) {
            const parsed = parseFontSize(inlineSize);
            if (parsed) return { raw: inlineSize, ...parsed };
        }

        // 向上走訪父節點找 inline style
        let parent = el.parentElement;
        while (parent) {
            const ps = parent.style?.fontSize;
            if (ps) {
                const parsed = parseFontSize(ps);
                if (parsed) return { raw: ps, ...parsed };
            }
            parent = parent.parentElement;
        }

        // Fallback：computed style（px，Odoo 預設行為）
        const computed = getComputedStyle(el).fontSize;
        if (computed) {
            const parsed = parseFontSize(computed);
            if (parsed) return { raw: computed, ...parsed };
        }
        return null;
    }

    /**
     * 對目前選取範圍套用字體大小。
     * @param {string} sizeStr - 如 "11pt"、"14px"、"11"（無單位時預設 pt）
     */
    _setFontSize(sizeStr) {
        let normalized = sizeStr.trim();
        if (!normalized) return;
        // 無單位數字預設加 pt
        if (/^\d+(\.\d+)?$/.test(normalized)) normalized += "pt";
        if (normalized === "pt") return;

        const docSel = window.getSelection();
        if (!docSel || docSel.rangeCount === 0) return;
        const range = docSel.getRangeAt(0);

        if (range.collapsed) {
            // 游標位置：直接對父元素設定（DOM mutation 自動被 MutationObserver 記錄）
            let el = range.startContainer;
            if (el.nodeType === 3) el = el.parentElement;
            if (el) el.style.fontSize = normalized;
        } else {
            // 有選取範圍：用 <span> 包裹選取文字
            const span = this.document.createElement("span");
            span.style.fontSize = normalized;
            try {
                range.surroundContents(span);
            } catch (e) {
                // 跨 block 選取（surroundContents 限制）：對 commonAncestor 設定
                const ancestor = range.commonAncestorContainer;
                const el = ancestor.nodeType === 3 ? ancestor.parentElement : ancestor;
                if (el) el.style.fontSize = normalized;
            }
        }
        this._updateDisplay();
    }
}
