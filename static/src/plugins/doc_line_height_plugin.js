/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { closestBlock } from "@html_editor/utils/blocks";
import { isVisibleTextNode } from "@html_editor/utils/dom_info";
import { withSequence } from "@html_editor/utils/resource";
import { toolbarButtonProps } from "@html_editor/main/toolbar/toolbar";
import { _t } from "@web/core/l10n/translation";
import { Component, useState, reactive } from "@odoo/owl";
import { Dropdown } from "@web/core/dropdown/dropdown";
import { DropdownItem } from "@web/core/dropdown/dropdown_item";

const LINE_HEIGHT_OPTIONS = [
    { name: "行高", value: "" },
    { name: "1.0", value: "1" },
    { name: "1.15", value: "1.15" },
    { name: "1.5", value: "1.5" },
    { name: "2.0", value: "2" },
    { name: "2.5", value: "2.5" },
    { name: "3.0", value: "3" },
];

/**
 * LineHeightSelector — 行高選擇器 OWL Component
 * 使用 Dropdown + DropdownItem（與 Odoo FontSelector 相同模式）
 * t-on-pointerdown.prevent 防止點選時 editor selection 遺失
 */
class LineHeightSelector extends Component {
    static template = "dobtor_doc_editor.LineHeightSelector";
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
 * DocLineHeightPlugin — 為 html_editor 浮動工具列增加行高選擇器
 *
 * 實作模式與 AlignPlugin 完全相同：
 * getTargetedNodes() → closestBlock() → block.style.lineHeight = value
 */
export class DocLineHeightPlugin extends Plugin {
    static id = "docLineHeight";
    static dependencies = ["selection"];

    setup() {
        // reactive 物件被 LineHeightSelector 的 useState 訂閱
        this.lineHeightState = reactive({ displayName: "行高" });
    }

    get resources() {
        return {
            toolbar_groups: [withSequence(26, { id: "docLineHeight" })],
            toolbar_items: [
                {
                    id: "lineHeight",
                    groupId: "docLineHeight",
                    title: _t("行高"),
                    Component: LineHeightSelector,
                    props: {
                        getItems: () => LINE_HEIGHT_OPTIONS,
                        getDisplay: () => this.lineHeightState,
                        onSelected: (item) => this._applyLineHeight(item.value),
                    },
                },
            ],
            // 游標移動時更新顯示的行高值
            selectionchange_handlers: [this._updateDisplay.bind(this)],
        };
    }

    _updateDisplay() {
        const sel = this.dependencies.selection.getEditableSelection?.();
        if (!sel?.anchorNode) {
            this.lineHeightState.displayName = "行高";
            return;
        }
        const block = closestBlock(sel.anchorNode);
        const lh = block?.style.lineHeight || "";
        const match = LINE_HEIGHT_OPTIONS.find((o) => o.value === lh);
        this.lineHeightState.displayName = match ? match.name : (lh || "行高");
    }

    _applyLineHeight(value) {
        const visitedBlocks = new Set();
        for (const node of this.dependencies.selection.getTargetedNodes()) {
            if (isVisibleTextNode(node)) {
                const block = closestBlock(node);
                if (block && !visitedBlocks.has(block) && block.isContentEditable) {
                    if (value) {
                        block.style.lineHeight = value;
                    } else {
                        block.style.removeProperty("line-height");
                    }
                    visitedBlocks.add(block);
                }
            }
        }
        // DOM mutation 自動被 MutationObserver 記錄，不需要手動 addStep
        this._updateDisplay();
    }
}
