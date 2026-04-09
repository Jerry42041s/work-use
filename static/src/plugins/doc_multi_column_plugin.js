/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { _t } from "@web/core/l10n/translation";
import { withSequence } from "@html_editor/utils/resource";

/**
 * DocMultiColumnPlugin — Section-Based 多欄排版。
 * 選取內容後可包裹為 2 或 3 欄 CSS column section。
 */
export class DocMultiColumnPlugin extends Plugin {
    static id = "docMultiColumn";
    static dependencies = ["history", "selection"];

    get resources() {
        return {
            powerbox_categories: [
                withSequence(45, {
                    id: "docLayout",
                    name: _t("排版"),
                }),
            ],
            user_commands: [
                {
                    id: "col1",
                    title: _t("1 欄（單欄）"),
                    description: _t("移除多欄排版，恢復單欄"),
                    icon: "fa-align-justify",
                    run: () => this._setColumns(1),
                },
                {
                    id: "col2",
                    title: _t("2 欄排版"),
                    description: _t("將選取或游標所在段落設為 2 欄"),
                    icon: "fa-columns",
                    run: () => this._setColumns(2),
                },
                {
                    id: "col3",
                    title: _t("3 欄排版"),
                    description: _t("將選取或游標所在段落設為 3 欄"),
                    icon: "fa-th",
                    run: () => this._setColumns(3),
                },
                {
                    id: "colBreak",
                    title: _t("插入換欄符號"),
                    description: _t("在多欄 section 內強制換欄"),
                    icon: "fa-pause",
                    run: () => this._insertColumnBreak(),
                },
            ],
            powerbox_items: [
                {
                    categoryId: "docLayout",
                    commandId: "col2",
                    title: _t("2 欄排版"),
                    icon: "fa-columns",
                    keywords: ["column", "欄", "多欄", "2col"],
                },
                {
                    categoryId: "docLayout",
                    commandId: "col3",
                    title: _t("3 欄排版"),
                    icon: "fa-th",
                    keywords: ["column", "欄", "多欄", "3col"],
                },
                {
                    categoryId: "docLayout",
                    commandId: "col1",
                    title: _t("恢復單欄"),
                    icon: "fa-align-justify",
                    keywords: ["single", "單欄", "1col"],
                },
                {
                    categoryId: "docLayout",
                    commandId: "colBreak",
                    title: _t("換欄"),
                    icon: "fa-pause",
                    keywords: ["column break", "換欄"],
                },
            ],
        };
    }

    _setColumns(columns) {
        const selection = this.document.getSelection();
        const existingSection = this._findColumnSection(selection?.anchorNode);

        if (existingSection) {
            if (columns === 1) {
                this._unwrapColumnSection(existingSection);
            } else {
                existingSection.dataset.columns = columns;
                existingSection.style.columnCount = columns;
            }
        } else {
            if (columns === 1) return;
            if (selection && !selection.isCollapsed) {
                this._wrapSelectionInColumnSection(selection, columns);
            } else {
                this._insertEmptyColumnSection(columns);
            }
        }

        if (this.shared.addStep) this.shared.addStep();
    }

    _insertEmptyColumnSection(columns) {
        const section = this.document.createElement("div");
        section.className = "doc-column-section";
        section.dataset.columns = columns;
        section.style.columnCount = columns;
        section.style.columnGap = "24px";
        section.innerHTML = "<p><br></p>";

        const sel = this.document.getSelection();
        if (sel?.anchorNode) {
            const block = this._findBlockParent(sel.anchorNode);
            if (block) {
                block.after(section);
                const range = this.document.createRange();
                range.setStart(section.querySelector("p"), 0);
                range.collapse(true);
                sel.removeAllRanges();
                sel.addRange(range);
                return;
            }
        }
        if (this.editable) this.editable.appendChild(section);
    }

    _wrapSelectionInColumnSection(selection, columns) {
        const range = selection.getRangeAt(0);
        const blocks = this._getSelectedTopLevelBlocks(range);
        if (!blocks.length) return;

        const section = this.document.createElement("div");
        section.className = "doc-column-section";
        section.dataset.columns = columns;
        section.style.columnCount = columns;
        section.style.columnGap = "24px";

        blocks[0].before(section);
        for (const block of blocks) section.appendChild(block);
    }

    _unwrapColumnSection(sectionEl) {
        const parent = sectionEl.parentNode;
        while (sectionEl.firstChild) {
            parent.insertBefore(sectionEl.firstChild, sectionEl);
        }
        parent.removeChild(sectionEl);
    }

    _insertColumnBreak() {
        const sel = this.document.getSelection();
        if (!sel || sel.rangeCount === 0) return;
        const section = this._findColumnSection(sel.anchorNode);
        if (!section) return;

        const breakEl = this.document.createElement("div");
        breakEl.className = "doc-column-break";
        breakEl.setAttribute("contenteditable", "false");

        const para = this.document.createElement("p");
        para.innerHTML = "<br>";

        const range = sel.getRangeAt(0);
        range.collapse(false);
        range.insertNode(para);
        range.insertNode(breakEl);

        const newRange = this.document.createRange();
        newRange.setStart(para, 0);
        newRange.collapse(true);
        sel.removeAllRanges();
        sel.addRange(newRange);

        if (this.shared.addStep) this.shared.addStep();
    }

    _findColumnSection(node) {
        let current = node;
        while (current) {
            if (current.nodeType === 1 && current.classList?.contains("doc-column-section")) {
                return current;
            }
            current = current.parentNode;
        }
        return null;
    }

    _findBlockParent(node) {
        const blockTags = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6",
                                    "DIV", "UL", "OL", "TABLE", "BLOCKQUOTE", "PRE"]);
        let current = node;
        while (current) {
            if (current.nodeType === 1 && blockTags.has(current.tagName)) return current;
            current = current.parentNode;
        }
        return null;
    }

    _getSelectedTopLevelBlocks(range) {
        const container = range.commonAncestorContainer;
        const editorEl = (container.nodeType === 1 ? container : container.parentElement)
            ?.closest("[contenteditable]") || this.editable;
        if (!editorEl) return [];
        return Array.from(editorEl.children).filter(child => range.intersectsNode(child));
    }
}
