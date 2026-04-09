/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { _t } from "@web/core/l10n/translation";
import { withSequence } from "@html_editor/utils/resource";

/**
 * DocPageFormatPlugin — 提供頁面相關的 Powerbox 指令：
 * - 插入分頁符號（Ctrl+Enter）
 * - 頁面格式切換（A4/A3/Letter）
 */
export class DocPageFormatPlugin extends Plugin {
    static id = "docPageFormat";
    static dependencies = ["history", "selection", "dom"];

    get resources() {
        return {
            powerbox_categories: [
                withSequence(40, {
                    id: "docPage",
                    name: _t("頁面"),
                }),
            ],
            user_commands: [
                {
                    id: "insertPageBreak",
                    title: _t("插入分頁符號"),
                    description: _t("強制從此處開始新頁"),
                    icon: "fa-file-o",
                    run: () => this.insertPageBreak(),
                },
                {
                    id: "setFormatA4",
                    title: _t("頁面格式：A4"),
                    description: _t("切換為 A4 (210 × 297mm)"),
                    icon: "fa-file",
                    run: () => this.setPageFormat("A4"),
                },
                {
                    id: "setFormatLetter",
                    title: _t("頁面格式：Letter"),
                    description: _t("切換為 Letter (216 × 279mm)"),
                    icon: "fa-file",
                    run: () => this.setPageFormat("letter"),
                },
            ],
            powerbox_items: [
                {
                    categoryId: "docPage",
                    commandId: "insertPageBreak",
                    title: _t("插入分頁符號"),
                    description: _t("Ctrl+Enter 或此指令"),
                    icon: "fa-file-o",
                    keywords: [_t("分頁"), "page break", "newpage"],
                },
                {
                    categoryId: "docPage",
                    commandId: "setFormatA4",
                    title: _t("A4 頁面"),
                    icon: "fa-file",
                },
                {
                    categoryId: "docPage",
                    commandId: "setFormatLetter",
                    title: _t("Letter 頁面"),
                    icon: "fa-file",
                },
            ],
        };
    }

    /**
     * 在游標位置插入分頁標記 div。
     */
    insertPageBreak() {
        const selection = this.shared.getEditableSelection
            ? this.shared.getEditableSelection()
            : this.document.getSelection();

        if (!selection) return;

        // 建立分頁 div
        const breakEl = this.document.createElement("div");
        breakEl.className = "doc-page-break";
        breakEl.setAttribute("contenteditable", "false");
        breakEl.setAttribute("data-page-break", "true");

        // 建立插入後的新段落（讓游標可以繼續輸入）
        const newPara = this.document.createElement("p");
        newPara.innerHTML = "<br>";

        try {
            // 取得選取範圍並插入
            if (selection.rangeCount > 0) {
                const range = selection.getRangeAt(0);
                range.collapse(false);
                range.insertNode(newPara);
                range.insertNode(breakEl);

                // 將游標移到新段落
                const newRange = this.document.createRange();
                newRange.setStart(newPara, 0);
                newRange.collapse(true);
                selection.removeAllRanges();
                selection.addRange(newRange);
            }
        } catch (e) {
            // fallback：在 editable 末尾插入
            if (this.editable) {
                this.editable.appendChild(breakEl);
                this.editable.appendChild(newPara);
            }
        }

        // 通知外層 DocEditor
        this.editable?.dispatchEvent(new CustomEvent("doc-insert-page-break", {
            bubbles: true,
        }));

        // 加入 undo history
        if (this.shared.addStep) {
            this.shared.addStep();
        }
    }

    /**
     * 切換頁面格式（通知外層 DocEditor 更新 state）。
     */
    setPageFormat(format) {
        this.editable?.dispatchEvent(new CustomEvent("doc-page-format-change", {
            detail: { format },
            bubbles: true,
        }));
    }
}
