/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { withSequence } from "@html_editor/utils/resource";
import { _t } from "@web/core/l10n/translation";

/**
 * DocTableMergePlugin — 表格儲存格合併/拆分
 *
 * 功能：
 * - 向右合併（mergeRight）：合併當前儲存格與右側儲存格
 * - 向下合併（mergeDown）：合併當前儲存格與下方儲存格
 * - 拆分儲存格（splitCell）：重設 colspan/rowspan 為 1
 *
 * 工具列按鈕僅在游標位於表格儲存格內時顯示（isAvailable）
 */
export class DocTableMergePlugin extends Plugin {
    static id = "docTableMerge";
    static dependencies = ["selection", "history"];

    get resources() {
        return {
            toolbar_groups: [withSequence(45, { id: "docTableMerge" })],
            user_commands: [
                {
                    id: "mergeCellRight",
                    title: _t("向右合併"),
                    icon: "fa-compress",
                    run: () => this._mergeRight(),
                },
                {
                    id: "mergeCellDown",
                    title: _t("向下合併"),
                    icon: "fa-compress",
                    run: () => this._mergeDown(),
                },
                {
                    id: "splitCell",
                    title: _t("拆分儲存格"),
                    icon: "fa-expand",
                    run: () => this._splitCell(),
                },
            ],
            toolbar_items: [
                {
                    id: "mergeCellRight",
                    groupId: "docTableMerge",
                    commandId: "mergeCellRight",
                    isAvailable: () => this._canMergeRight(),
                },
                {
                    id: "mergeCellDown",
                    groupId: "docTableMerge",
                    commandId: "mergeCellDown",
                    isAvailable: () => this._canMergeDown(),
                },
                {
                    id: "splitCell",
                    groupId: "docTableMerge",
                    commandId: "splitCell",
                    isAvailable: () => this._isInMergedCell(),
                },
            ],
        };
    }

    // ── 取得游標所在的 <td> 或 <th> ──
    _getActiveTd() {
        const sel = this.dependencies.selection.getEditableSelection?.();
        if (!sel) return null;
        const node = sel.anchorNode;
        if (!node) return null;
        const el = node.nodeType === 3 ? node.parentElement : node;
        return el?.closest?.("td, th") || null;
    }

    _isInMergedCell() {
        const td = this._getActiveTd();
        if (!td) return false;
        const colspan = parseInt(td.getAttribute("colspan") || "1");
        const rowspan = parseInt(td.getAttribute("rowspan") || "1");
        return colspan > 1 || rowspan > 1;
    }

    _canMergeRight() {
        const td = this._getActiveTd();
        if (!td) return false;
        return !!td.nextElementSibling;
    }

    _canMergeDown() {
        const td = this._getActiveTd();
        if (!td) return false;
        const table = td.closest("table");
        if (!table) return false;
        const rowIndex = td.parentElement.rowIndex;
        const rowspan = parseInt(td.getAttribute("rowspan") || "1");
        return rowIndex + rowspan < table.rows.length;
    }

    // ── 向右合併：合併右側儲存格（colspan++） ──
    _mergeRight() {
        const td = this._getActiveTd();
        if (!td) return;
        const next = td.nextElementSibling;
        if (!next) return;

        const colspan = parseInt(td.getAttribute("colspan") || "1");
        const nextColspan = parseInt(next.getAttribute("colspan") || "1");
        td.setAttribute("colspan", colspan + nextColspan);

        // 將右側內容移入（以換行分隔）
        if (next.innerHTML.trim()) {
            td.innerHTML += "<br>" + next.innerHTML;
        }
        next.remove();
        this.dependencies.history.addStep();
    }

    // ── 向下合併：合併下方儲存格（rowspan++） ──
    _mergeDown() {
        const td = this._getActiveTd();
        if (!td) return;
        const table = td.closest("table");
        if (!table) return;

        const rowspan = parseInt(td.getAttribute("rowspan") || "1");
        const rowIndex = td.parentElement.rowIndex;
        const nextRow = table.rows[rowIndex + rowspan];
        if (!nextRow) return;

        // 計算當前儲存格的欄位索引
        let colIndex = 0;
        const cells = td.parentElement.cells;
        for (let i = 0; i < td.cellIndex; i++) {
            colIndex += parseInt(cells[i].getAttribute("colspan") || "1");
        }

        // 在下一行找到對應欄位索引的儲存格
        let accCol = 0;
        let targetTd = null;
        for (const cell of nextRow.cells) {
            if (accCol === colIndex) {
                targetTd = cell;
                break;
            }
            accCol += parseInt(cell.getAttribute("colspan") || "1");
        }
        if (!targetTd) return;

        const nextRowspan = parseInt(targetTd.getAttribute("rowspan") || "1");
        td.setAttribute("rowspan", rowspan + nextRowspan);

        if (targetTd.innerHTML.trim()) {
            td.innerHTML += "<br>" + targetTd.innerHTML;
        }
        targetTd.remove();
        this.dependencies.history.addStep();
    }

    // ── 拆分儲存格：移除 colspan/rowspan ──
    _splitCell() {
        const td = this._getActiveTd();
        if (!td) return;
        td.removeAttribute("colspan");
        td.removeAttribute("rowspan");
        this.dependencies.history.addStep();
    }
}
