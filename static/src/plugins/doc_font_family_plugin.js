/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { formatsSpecs } from "@html_editor/utils/formatting";
import { withSequence } from "@html_editor/utils/resource";
import { toolbarButtonProps } from "@html_editor/main/toolbar/toolbar";
import { _t } from "@web/core/l10n/translation";
import { Component, useState, reactive } from "@odoo/owl";
import { Dropdown } from "@web/core/dropdown/dropdown";
import { DropdownItem } from "@web/core/dropdown/dropdown_item";

// ── 注入 fontFamily 格式規格到 html_editor formatsSpecs ──
if (!formatsSpecs.fontFamily) {
    formatsSpecs.fontFamily = {
        isFormatted: (node) => Boolean(node.style?.["font-family"]),
        hasStyle: (node) => Boolean(node.style?.["font-family"]),
        addStyle: (node, props) => {
            node.style["font-family"] = props.fontFamily;
        },
        removeStyle: (node) => {
            node.style.removeProperty("font-family");
        },
    };
}

// 僅保留跨平台可靠的系統字型 + Windows 常見中文字型
const FONT_FAMILIES = [
    { name: "預設字型", value: "" },
    { name: "Arial", value: "Arial, sans-serif" },
    { name: "Times New Roman", value: "'Times New Roman', Times, serif" },
    { name: "Courier New", value: "'Courier New', Courier, monospace" },
    { name: "Georgia", value: "Georgia, serif" },
    { name: "Verdana", value: "Verdana, Geneva, sans-serif" },
    { name: "微軟正黑體", value: "'Microsoft JhengHei', '微軟正黑體', sans-serif" },
    { name: "新細明體", value: "'PMingLiU', '新細明體', serif" },
    { name: "標楷體", value: "'DFKai-SB', '標楷體', cursive" },
];

/**
 * FontFamilySelector — 字型選擇器 OWL Component
 * 使用 Dropdown + DropdownItem（與 Odoo FontSelector 相同模式）
 * t-on-pointerdown.prevent 防止點選時 editor selection 遺失
 */
class FontFamilySelector extends Component {
    static template = "dobtor_doc_editor.FontFamilySelector";
    static components = { Dropdown, DropdownItem };
    static props = {
        getItems: Function,
        getDisplay: Function,
        onSelected: Function,
        ...toolbarButtonProps,
    };

    setup() {
        this.items = this.props.getItems();
        // getDisplay() 回傳 reactive 物件，useState 使其在此 Component 內可被觀察
        this.state = useState(this.props.getDisplay());
    }

    onSelected(item) {
        this.props.onSelected(item);
    }
}

/**
 * DocFontFamilyPlugin — 為 html_editor 浮動工具列增加字型選擇器
 * - selectionchange_handlers 在游標移動時更新顯示名稱
 */
export class DocFontFamilyPlugin extends Plugin {
    static id = "docFontFamily";
    static dependencies = ["format", "selection"];

    setup() {
        // reactive 物件被 FontFamilySelector 的 useState 訂閱
        // 直接 mutate 此物件即可觸發 Component 重渲染
        this.fontFamilyState = reactive({ displayName: "字型" });
    }

    get resources() {
        return {
            toolbar_groups: [withSequence(4, { id: "docFontFamily" })],
            toolbar_items: [
                {
                    id: "fontFamily",
                    groupId: "docFontFamily",
                    title: _t("字型"),
                    Component: FontFamilySelector,
                    props: {
                        getItems: () => FONT_FAMILIES,
                        getDisplay: () => this.fontFamilyState,
                        onSelected: (item) => this._applyFontFamily(item.value),
                    },
                },
            ],
            // 游標移動時更新顯示的字型名稱
            selectionchange_handlers: [this._updateDisplay.bind(this)],
        };
    }

    _updateDisplay() {
        const sel = this.dependencies.selection.getEditableSelection?.();
        if (!sel?.anchorNode) {
            this.fontFamilyState.displayName = "字型";
            return;
        }
        const el = sel.anchorNode.nodeType === 3
            ? sel.anchorNode.parentElement
            : sel.anchorNode;
        const ff = el ? getComputedStyle(el).fontFamily : "";
        // 比對已知字型清單（移除引號後比對第一個字型名稱）
        const normalize = (s) => s.toLowerCase().replace(/['"]/g, "").trim();
        const match = FONT_FAMILIES.find((f) => {
            if (!f.value) return false;
            const first = normalize(f.value.split(",")[0]);
            return normalize(ff).includes(first);
        });
        this.fontFamilyState.displayName = match ? match.name : "字型";
    }

    _applyFontFamily(fontFamily) {
        if (!fontFamily) {
            this.dependencies.format.formatSelection("fontFamily", { applyStyle: false });
        } else {
            this.dependencies.format.formatSelection("fontFamily", {
                applyStyle: true,
                formatProps: { fontFamily },
            });
        }
        this._updateDisplay();
    }
}
