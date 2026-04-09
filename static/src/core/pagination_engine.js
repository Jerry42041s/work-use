/** @odoo-module **/

/**
 * PaginationEngine — 基於 DOM element 高度的精確分頁引擎。
 *
 * 設計原則：
 * 1. 遍歷 contentEditable 內的頂層 block elements
 * 2. 根據每個 element 的 offsetHeight 決定分頁位置
 * 3. 表格支援 row-level 斷頁
 * 4. 支援孤行/寡行控制
 * 5. 回傳分頁斷點清單（供外層渲染視覺分頁線）
 */
export class PaginationEngine {
    /**
     * @param {Object} options
     * @param {number} options.pageContentHeight - 單頁可用內容高度（px）
     * @param {number} options.orphanMinLines - 孤行控制（預設 2）
     * @param {number} options.widowMinLines  - 寡行控制（預設 2）
     */
    constructor(options) {
        this.pageContentHeight = options.pageContentHeight;
        this.orphanMinLines = options.orphanMinLines ?? 2;
        this.widowMinLines = options.widowMinLines ?? 2;
        this._lineHeightCache = new Map();
    }

    /**
     * 計算分頁斷點。
     * @param {HTMLElement} contentEl - 編輯區域的 root element
     * @returns {{ pageIndex: number, topPx: number }[]} 每頁的起始位置（相對於 contentEl top）
     */
    computePageBreaks(contentEl) {
        if (!contentEl) return [];

        const blocks = Array.from(contentEl.children);
        const pages = [{ pageIndex: 0, topPx: 0 }];
        let usedHeight = 0;
        let pageIndex = 0;
        const contentTop = contentEl.getBoundingClientRect().top;

        for (const block of blocks) {
            // ── 跳過 DOM 分頁間隔元素（由 DocPageLayout 插入，不算入內容高度）──
            if (block.classList.contains('doc-page-separator')) continue;

            // ── Hard Page Break ──
            if (
                block.classList.contains("doc-page-break-hard") ||
                block.dataset?.pageBreak === "true"
            ) {
                pageIndex++;
                const blockTop = block.getBoundingClientRect().top - contentTop;
                pages.push({ pageIndex, topPx: blockTop, blockRef: block });
                usedHeight = 0;
                continue;
            }

            const blockHeight = block.getBoundingClientRect().height;

            if (blockHeight <= 0) continue;

            // ── Column Section ──
            if (block.classList.contains("doc-column-section")) {
                const colHeight = this._measureColumnSectionHeight(block);
                if (usedHeight + colHeight > this.pageContentHeight && usedHeight > 0) {
                    pageIndex++;
                    const blockTop = block.getBoundingClientRect().top - contentTop;
                    pages.push({ pageIndex, topPx: blockTop });
                    usedHeight = colHeight;
                } else {
                    usedHeight += colHeight;
                }
                continue;
            }

            if (usedHeight + blockHeight <= this.pageContentHeight) {
                usedHeight += blockHeight;
            } else if (block.tagName === "TABLE") {
                // 表格：嘗試 row-level 斷頁
                const breakPoints = this._paginateTable(block, usedHeight, contentTop, pageIndex);
                for (const bp of breakPoints) {
                    pages.push(bp);
                    pageIndex = bp.pageIndex;
                }
                usedHeight = this._getLastRowsHeight(block);
            } else if (this._isParagraph(block)) {
                const decision = this._decideParagraphBreak(block, blockHeight, usedHeight);
                if (decision === "push") {
                    pageIndex++;
                    const blockTop = block.getBoundingClientRect().top - contentTop;
                    pages.push({ pageIndex, topPx: blockTop, blockRef: block });
                    usedHeight = blockHeight;
                } else {
                    // 允許段落跨頁：在目前位置換頁
                    if (usedHeight > 0) {
                        pageIndex++;
                        pages.push({ pageIndex, topPx: usedHeight, blockRef: block });
                    }
                    usedHeight = blockHeight % this.pageContentHeight;
                }
            } else {
                // 圖片或其他元素：整個推到下一頁
                if (usedHeight > 0) {
                    pageIndex++;
                    const blockTop = block.getBoundingClientRect().top - contentTop;
                    pages.push({ pageIndex, topPx: blockTop, blockRef: block });
                }
                usedHeight = blockHeight;
            }
        }

        return pages;
    }

    _paginateTable(tableEl, usedHeight, contentTop, currentPageIndex) {
        const rows = Array.from(tableEl.querySelectorAll(":scope > thead > tr, :scope > tbody > tr"));
        const breakPoints = [];
        let accumulated = usedHeight;
        let pageIndex = currentPageIndex;

        for (const tr of rows) {
            const rowHeight = tr.getBoundingClientRect().height;
            if (accumulated + rowHeight > this.pageContentHeight && accumulated > 0) {
                pageIndex++;
                const rowTop = tr.getBoundingClientRect().top - contentTop;
                breakPoints.push({ pageIndex, topPx: rowTop, blockRef: tr });
                accumulated = rowHeight;
            } else {
                accumulated += rowHeight;
            }
        }
        return breakPoints;
    }

    _getLastRowsHeight(tableEl) {
        // 取得最後一頁剩餘的表格高度（近似）
        const rows = Array.from(tableEl.querySelectorAll(":scope > tbody > tr"));
        if (!rows.length) return 0;
        let h = 0;
        for (const tr of [...rows].reverse()) {
            h += tr.getBoundingClientRect().height;
            if (h > this.pageContentHeight) break;
        }
        return Math.min(h, this.pageContentHeight);
    }

    _decideParagraphBreak(block, blockHeight, usedHeight) {
        const lineHeight = this._estimateLineHeight(block);
        if (lineHeight <= 0) return "push";
        const linesThatFit = Math.floor((this.pageContentHeight - usedHeight) / lineHeight);
        if (linesThatFit < this.orphanMinLines) return "push";
        const totalLines = Math.round(blockHeight / lineHeight);
        const linesOnNextPage = totalLines - linesThatFit;
        if (linesOnNextPage > 0 && linesOnNextPage < this.widowMinLines) return "push";
        return "split";
    }

    _estimateLineHeight(el) {
        const tag = el.tagName;
        if (this._lineHeightCache.has(tag)) return this._lineHeightCache.get(tag);
        const computed = window.getComputedStyle(el);
        const lh = parseFloat(computed.lineHeight) || parseFloat(computed.fontSize) * 1.5 || 24;
        this._lineHeightCache.set(tag, lh);
        return lh;
    }

    _measureColumnSectionHeight(sectionEl) {
        // CSS column 容器的 offsetHeight 就是視覺高度
        return sectionEl.getBoundingClientRect().height;
    }

    _isParagraph(el) {
        return ["P", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "PRE"].includes(el.tagName);
    }
}
