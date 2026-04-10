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
            if (block.classList.contains('doc-page-gap')) continue;

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


// ── Debounce 工具函式 ────────────────────────────────────────────────────────
function _debounce(fn, wait) {
    let timer;
    return function (...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), wait);
    };
}


/**
 * PaginationController — 真實 DOM 多頁分頁控制器
 *
 * 架構說明：
 * - 維護單一 Wysiwyg contenteditable 內的多個 .doc-page-sheet 子 div
 * - 每個 .doc-page-sheet 代表一頁，視覺上呈現 A4 白紙效果
 * - .doc-page-gap 為不可編輯的頁間灰色間隔
 * - ResizeObserver（250ms debounce）監聽內容溢出，自動搬移 block 節點
 * - 所有節點搬移前後保存/恢復游標（Selection/Range Restoration）
 * - ArrowDown/Up keydown 攔截，跳過 .doc-page-gap 防止游標卡住
 *
 * @param {Object} options
 * @param {HTMLElement} options.container      - Wysiwyg 的 contenteditable 根元素
 * @param {number}      options.pageWidth      - 頁面寬度（px）
 * @param {number}      options.pageHeight     - 頁面總高度（px，含邊距）
 * @param {number}      options.marginTop      - 上邊距（px）
 * @param {number}      options.marginBottom   - 下邊距（px）
 * @param {number}      options.marginLeft     - 左邊距（px）
 * @param {number}      options.marginRight    - 右邊距（px）
 */
export class PaginationController {
    constructor({ container, pageWidth, pageHeight, marginTop, marginBottom, marginLeft, marginRight }) {
        this.container = container;
        this.pageWidth = pageWidth;
        this.pageHeight = pageHeight;
        this.marginTop = marginTop;
        this.marginBottom = marginBottom;
        this.marginLeft = marginLeft;
        this.marginRight = marginRight;

        // 可用內容高度（扣除邊距）
        this.pageContentHeight = pageHeight - marginTop - marginBottom;

        this._pages = [];       // Array<HTMLElement>（.doc-page-sheet）
        this._destroyed = false;
        this.onPagesChange = null;  // 可選回呼：(pageCount: number) => void

        // Debounce 版 _redistribute（ResizeObserver 使用）
        this._redistributeDebounced = _debounce(
            (fromIdx) => this._redistribute(fromIdx),
            250
        );

        // keydown handler（跳過 gap）
        this._boundSkipGap = this._skipGapOnArrow.bind(this);
    }

    // ── 初始化 ───────────────────────────────────────────────────────────────

    /**
     * 包裹 container 現有的子節點至第一個 .doc-page-sheet，然後開始分頁管理。
     * 在 Wysiwyg 已完成掛載後呼叫（不清空現有 DOM）。
     */
    wrapExistingContent() {
        // 若已有 .doc-page-sheet，直接重用
        const existingSheets = Array.from(
            this.container.querySelectorAll(':scope > .doc-page-sheet')
        );
        if (existingSheets.length > 0) {
            this._pages = existingSheets;
        } else {
            // 建立第一頁，並搬入所有現有子節點
            const firstSheet = this._createPageSheet(0);
            const existing = Array.from(this.container.childNodes).filter(n => {
                if (n.nodeType === Node.ELEMENT_NODE) {
                    return !n.classList.contains('doc-page-sheet') &&
                           !n.classList.contains('doc-page-gap');
                }
                return true;
            });
            for (const child of existing) {
                firstSheet.appendChild(child);
            }
            // 確保至少有一個空段落
            if (firstSheet.children.length === 0) {
                const p = document.createElement('p');
                p.appendChild(document.createElement('br'));
                firstSheet.appendChild(p);
            }
            this.container.appendChild(firstSheet);
            this._pages = [firstSheet];
        }

        // 執行一次分配
        this._redistribute(0);

        // 綁定 ResizeObserver
        this._resizeObserver = new ResizeObserver(() => {
            if (!this._destroyed) this._redistributeDebounced(0);
        });
        this._resizeObserver.observe(this.container);

        // 綁定 MutationObserver：偵測 history.reset 一次清空所有 page sheets 的情境。
        // 正常分頁操作只移除尾部空 sheet，不會清空全部；history.reset 才會全清。
        this._mutationObserver = new MutationObserver((mutations) => {
            if (this._destroyed) return;
            for (const m of mutations) {
                const removedSheets = Array.from(m.removedNodes)
                    .filter(n => n.classList?.contains('doc-page-sheet'));
                if (removedSheets.length === 0) continue;
                const remaining = this.container.querySelectorAll(':scope > .doc-page-sheet');
                if (remaining.length === 0) {
                    this._pages = [];
                    // setTimeout(0)：讓 history.reset 的 innerHTML 替換完成後再重包裹
                    setTimeout(() => {
                        if (!this._destroyed) this._rewrapFlatContent();
                    }, 0);
                    return;
                }
            }
        });
        this._mutationObserver.observe(this.container, { childList: true });

        // 綁定 Arrow Key handler
        this.container.addEventListener('keydown', this._boundSkipGap);
    }

    /**
     * 初始化：清空 container，用提供的 HTML 建立第一頁並分配內容。
     * @param {string} htmlContent - 初始 HTML 字串
     */
    init(htmlContent) {
        // 清空現有內容
        this.container.innerHTML = '';
        this._pages = [];

        // 建立第一頁
        const firstSheet = this._createPageSheet(0);
        this.container.appendChild(firstSheet);
        this._pages.push(firstSheet);

        // 寫入初始內容
        const contentArea = this._contentArea(firstSheet);
        contentArea.innerHTML = htmlContent || '<p><br></p>';

        // 立即執行一次分配（非 debounce）
        this._redistribute(0);

        // 綁定 ResizeObserver
        this._resizeObserver = new ResizeObserver(() => {
            if (!this._destroyed) this._redistributeDebounced(0);
        });
        this._resizeObserver.observe(this.container);

        // 綁定 MutationObserver（同 wrapExistingContent）
        this._mutationObserver = new MutationObserver((mutations) => {
            if (this._destroyed) return;
            for (const m of mutations) {
                const removedSheets = Array.from(m.removedNodes)
                    .filter(n => n.classList?.contains('doc-page-sheet'));
                if (removedSheets.length === 0) continue;
                const remaining = this.container.querySelectorAll(':scope > .doc-page-sheet');
                if (remaining.length === 0) {
                    this._pages = [];
                    setTimeout(() => {
                        if (!this._destroyed) this._rewrapFlatContent();
                    }, 0);
                    return;
                }
            }
        });
        this._mutationObserver.observe(this.container, { childList: true });

        // 綁定 Arrow Key handler
        this.container.addEventListener('keydown', this._boundSkipGap);
    }

    /**
     * 銷毀：解除所有監聽器。
     */
    destroy() {
        this._destroyed = true;
        if (this._resizeObserver) {
            this._resizeObserver.disconnect();
            this._resizeObserver = null;
        }
        if (this._mutationObserver) {
            this._mutationObserver.disconnect();
            this._mutationObserver = null;
        }
        this.container.removeEventListener('keydown', this._boundSkipGap);
    }

    // ── 序列化 ───────────────────────────────────────────────────────────────

    /**
     * 取得所有頁面的內容，合併為一個 HTML 字串（不含 layout 元素）。
     * @returns {string}
     */
    serialize() {
        return this._pages
            .map(sheet => this._contentArea(sheet).innerHTML)
            .join('');
    }

    // ── 分頁分配 ─────────────────────────────────────────────────────────────

    /**
     * 從指定頁開始，重新分配所有頁面的內容。
     * @param {number} fromPageIdx
     */
    _redistribute(fromPageIdx) {
        if (this._destroyed || !this._pages.length) return;

        // ── 自我修復：若首頁已脫離 container（如 history.reset 後），重新包裹平面內容 ──
        if (this._pages[0] && !this.container.contains(this._pages[0])) {
            this._pages = [];
            this._rewrapFlatContent();
            return;  // _rewrapFlatContent 內部會觸發新一輪 _redistribute
        }

        const startIdx = Math.max(0, fromPageIdx);

        // 從 startIdx 頁開始，逐頁處理溢出與下溢
        for (let i = startIdx; i < this._pages.length; i++) {
            // 1. 將溢出的 blocks 推到下一頁
            this._drainOverflow(i);
            // 2. 嘗試從下一頁拉回 blocks（可能下溢）
            this._fillUnderflow(i);
        }

        // 移除尾端空頁
        this._pruneEmptyPages();

        // 通知頁數變更
        if (typeof this.onPagesChange === 'function') {
            this.onPagesChange(this._pages.length);
        }
    }

    /**
     * 持續將頁面 pageIdx 的溢出 blocks 推到下一頁，直到不再溢出。
     *
     * 無限迴圈防護：若只剩一個無法再拆分的 block（如單行表格或單一段落），
     * 強制 break，寧可視覺溢出也不當機。
     *
     * TABLE 特例：呼叫 _splitTableOverflow 進行列級拆分，避免整表搬移造成
     * 第一頁空白或單頁表格引發死迴圈。
     */
    _drainOverflow(pageIdx) {
        const sheet = this._pages[pageIdx];
        if (!sheet) return;

        const contentArea = this._contentArea(sheet);
        let iterations = 0;
        const MAX_ITER = 200;

        while (this._isOverflowing(contentArea) && iterations++ < MAX_ITER) {
            const lastBlock = this._lastBlock(contentArea);
            if (!lastBlock) break;

            // ── 無限迴圈防護：計算可見 block 數量 ──
            const visibleBlockCount = Array.from(contentArea.children)
                .filter(n => !n.classList.contains('doc-page-gap')).length;
            const isOnlyBlock = visibleBlockCount <= 1;

            // ── TABLE：嘗試列級拆分 ──
            if (lastBlock.tagName === 'TABLE') {
                const moved = this._splitTableOverflow(pageIdx, contentArea, lastBlock);
                if (!moved) {
                    // 無法列級拆分（只剩 1 行或無 tbody）：
                    // 若還有其他 blocks，搬整個 table；若唯一 block，強制退出。
                    if (!isOnlyBlock) {
                        const saved = this._saveSelection();
                        this._ensureNextPage(pageIdx);
                        const nextCA = this._contentArea(this._pages[pageIdx + 1]);
                        contentArea.removeChild(lastBlock);
                        nextCA.insertBefore(lastBlock, nextCA.firstChild);
                        this._restoreSelection(saved);
                    } else {
                        break;  // 唯一 block 且無法拆，強制退出
                    }
                }
                // _splitTableOverflow 內部已執行到不溢出，continue 重新檢查
                continue;
            }

            // ── 一般 block：若是唯一 block，強制退出（寧可視覺溢出）──
            if (isOnlyBlock) break;

            // ── 一般 block：搬移至下一頁開頭 ──
            const saved = this._saveSelection();
            this._ensureNextPage(pageIdx);
            const nextContentArea = this._contentArea(this._pages[pageIdx + 1]);
            contentArea.removeChild(lastBlock);
            nextContentArea.insertBefore(lastBlock, nextContentArea.firstChild);
            this._restoreSelection(saved);
        }
    }

    /**
     * 嘗試將下一頁的第一個 block 拉回本頁（若本頁有空間且不會造成溢出）。
     *
     * TABLE 特例：若本頁末尾是 TABLE 且下一頁開頭也是 TABLE（分割的延續表格），
     * 呼叫 _mergeTableUnderflow 以列為單位回拉，而非搬整個 table。
     */
    _fillUnderflow(pageIdx) {
        if (pageIdx + 1 >= this._pages.length) return;

        const contentArea = this._contentArea(this._pages[pageIdx]);
        const nextContentArea = this._contentArea(this._pages[pageIdx + 1]);

        let iterations = 0;
        const MAX_ITER = 100;

        while (iterations++ < MAX_ITER) {
            const firstNextBlock = nextContentArea.firstElementChild;
            if (!firstNextBlock) break;

            const lastCurrentBlock = this._lastBlock(contentArea);

            // ── 表格行回拉合併 ──
            if (lastCurrentBlock && lastCurrentBlock.tagName === 'TABLE' &&
                    firstNextBlock.tagName === 'TABLE') {
                this._mergeTableUnderflow(pageIdx);
                break;
            }

            // ── 一般 block ──
            const saved = this._saveSelection();
            contentArea.appendChild(firstNextBlock);

            if (this._isOverflowing(contentArea)) {
                // 放不下：退回
                contentArea.removeChild(firstNextBlock);
                nextContentArea.insertBefore(firstNextBlock, nextContentArea.firstChild);
                this._restoreSelection(saved);
                break;
            }

            this._restoreSelection(saved);
        }
    }

    // ── 表格列級拆分 ─────────────────────────────────────────────────────────

    /**
     * 將溢出的表格從底部逐列搬移至下一頁的延續表格，直到本頁不再溢出。
     *
     * @param {number}      pageIdx     - 目前頁面索引
     * @param {HTMLElement} contentArea - 目前頁面的內容區域
     * @param {HTMLElement} tableEl     - 溢出的 TABLE 元素
     * @returns {boolean} 是否成功搬移了至少一行
     */
    _splitTableOverflow(pageIdx, contentArea, tableEl) {
        const tbody = tableEl.querySelector(':scope > tbody');
        if (!tbody) return false;  // 無 tbody，無法列級拆分

        const rows = Array.from(tbody.querySelectorAll(':scope > tr'));
        if (rows.length <= 1) return false;  // 只剩 1 行，無法繼續拆分

        // 確保下一頁存在
        this._ensureNextPage(pageIdx);
        const nextContentArea = this._contentArea(this._pages[pageIdx + 1]);

        // 取得或建立下一頁的延續表格
        let contTable;
        const firstNext = nextContentArea.firstElementChild;
        if (firstNext && firstNext.tagName === 'TABLE') {
            contTable = firstNext;
        } else {
            contTable = this._createContinuationTable(tableEl);
            nextContentArea.insertBefore(contTable, nextContentArea.firstChild);
        }

        let contTbody = contTable.querySelector(':scope > tbody');
        if (!contTbody) {
            contTbody = document.createElement('tbody');
            contTable.appendChild(contTbody);
        }

        let moved = false;
        const saved = this._saveSelection();

        // 從底部逐列搬移，直到不再溢出或只剩 1 行
        while (this._isOverflowing(contentArea)) {
            const currentRows = Array.from(tbody.querySelectorAll(':scope > tr'));
            if (currentRows.length <= 1) break;

            const lastRow = currentRows[currentRows.length - 1];
            tbody.removeChild(lastRow);
            contTbody.insertBefore(lastRow, contTbody.firstChild);
            moved = true;
        }

        this._restoreSelection(saved);
        return moved;
    }

    /**
     * 建立延續表格（複製 colgroup + thead 結構，tbody 為空）。
     * @param {HTMLElement} sourceTable - 原始表格
     * @returns {HTMLElement}
     */
    _createContinuationTable(sourceTable) {
        const newTable = document.createElement('table');
        // 複製屬性（class、style、border 等），跳過 id 避免重複
        for (const attr of sourceTable.attributes) {
            if (attr.name === 'id') continue;
            newTable.setAttribute(attr.name, attr.value);
        }
        // 複製 colgroup（保留欄寬設定）
        const colgroup = sourceTable.querySelector(':scope > colgroup');
        if (colgroup) newTable.appendChild(colgroup.cloneNode(true));
        // 複製 thead（保留欄標題）
        const thead = sourceTable.querySelector(':scope > thead');
        if (thead) newTable.appendChild(thead.cloneNode(true));
        // 空 tbody 接收搬來的 <tr>
        newTable.appendChild(document.createElement('tbody'));
        return newTable;
    }

    /**
     * 將下一頁延續表格的列回拉至本頁表格（下溢補償）。
     * 若延續表格的 tbody 被清空，移除整個延續表格。
     * @param {number} pageIdx
     */
    _mergeTableUnderflow(pageIdx) {
        if (pageIdx + 1 >= this._pages.length) return;

        const contentArea = this._contentArea(this._pages[pageIdx]);
        const nextContentArea = this._contentArea(this._pages[pageIdx + 1]);

        const lastBlock = this._lastBlock(contentArea);
        const firstNextBlock = nextContentArea.firstElementChild;
        if (!lastBlock || !firstNextBlock) return;
        if (lastBlock.tagName !== 'TABLE' || firstNextBlock.tagName !== 'TABLE') return;

        const tbody = lastBlock.querySelector(':scope > tbody');
        const nextTbody = firstNextBlock.querySelector(':scope > tbody');
        if (!tbody || !nextTbody) return;

        const saved = this._saveSelection();

        // 逐列從下一頁的延續表格拉回本頁
        while (nextTbody.firstElementChild) {
            const firstRow = nextTbody.firstElementChild;
            tbody.appendChild(firstRow);

            if (this._isOverflowing(contentArea)) {
                // 放不下：退回
                tbody.removeChild(firstRow);
                nextTbody.insertBefore(firstRow, nextTbody.firstChild);
                break;
            }
        }

        // 若延續表格的 tbody 已清空，移除整個延續表格
        if (nextTbody.children.length === 0) {
            nextContentArea.removeChild(firstNextBlock);
        }

        this._restoreSelection(saved);
    }

    /**
     * 確保 pageIdx + 1 的頁面存在，若不存在則建立。
     */
    _ensureNextPage(pageIdx) {
        if (pageIdx + 1 < this._pages.length) return;

        const newSheet = this._createPageSheet(pageIdx + 1);
        const gap = this._createPageGap();

        // 在 container 尾端加入 gap + 新頁
        this.container.appendChild(gap);
        this.container.appendChild(newSheet);
        this._pages.push(newSheet);
    }

    /**
     * 移除尾端空頁（最後一頁若無內容且非唯一頁，移除）。
     */
    _pruneEmptyPages() {
        while (this._pages.length > 1) {
            const lastSheet = this._pages[this._pages.length - 1];
            const contentArea = this._contentArea(lastSheet);

            if (this._isEmpty(contentArea)) {
                // 移除最後一頁及其前面的 gap
                const lastIdx = this._pages.length - 1;
                this._pages.pop();

                // 移除 DOM：先移 gap（lastSheet 的前兄弟），再移 sheet
                const prevSibling = lastSheet.previousSibling;
                if (prevSibling && prevSibling.classList &&
                    prevSibling.classList.contains('doc-page-gap')) {
                    this.container.removeChild(prevSibling);
                }
                this.container.removeChild(lastSheet);
            } else {
                break;
            }
        }
    }

    // ── DOM 工廠 ─────────────────────────────────────────────────────────────

    /**
     * 平面內容自我修復：將 container 直接子節點包進新 .doc-page-sheet。
     * 用於 history.reset 或外部 innerHTML 替換後的結構恢復。
     * 不重設 ResizeObserver / keydown（仍由原有監聽器服務）。
     */
    _rewrapFlatContent() {
        const sheet = this._createPageSheet(0);
        const children = Array.from(this.container.childNodes).filter(n => {
            if (n.nodeType === Node.ELEMENT_NODE) {
                return !n.classList.contains('doc-page-sheet') &&
                       !n.classList.contains('doc-page-gap');
            }
            return true;
        });
        for (const child of children) {
            sheet.appendChild(child);
        }
        if (sheet.children.length === 0) {
            const p = document.createElement('p');
            p.appendChild(document.createElement('br'));
            sheet.appendChild(p);
        }
        this.container.appendChild(sheet);
        this._pages = [sheet];
        this._redistribute(0);
    }

    /**
     * 建立 .doc-page-sheet 容器。
     * 包含一個內部 .doc-page-content 作為實際內容區域。
     */
    _createPageSheet(idx) {
        const sheet = document.createElement('div');
        sheet.className = 'doc-page-sheet';
        sheet.dataset.page = String(idx);
        sheet.style.cssText = [
            `width: ${this.pageWidth}px`,
            `min-height: ${this.pageHeight}px`,
            `padding: ${this.marginTop}px ${this.marginRight}px ${this.marginBottom}px ${this.marginLeft}px`,
            `box-sizing: border-box`,
            `background: #fff`,
            `box-shadow: 0 2px 8px rgba(0,0,0,0.25)`,
            `margin: 0 auto`,
            `position: relative`,
        ].join('; ');

        // 內容區域（block 節點直接放在 sheet 內，不額外包一層）
        // sheet 本身即 contentArea
        return sheet;
    }

    /**
     * 建立頁間 .doc-page-gap（不可編輯的灰色間隔）。
     */
    _createPageGap() {
        const gap = document.createElement('div');
        gap.className = 'doc-page-gap';
        gap.contentEditable = 'false';
        gap.dataset.pageGap = 'true';
        gap.style.cssText = [
            `height: 32px`,
            `background: transparent`,
            `user-select: none`,
            `pointer-events: none`,
            `display: block`,
        ].join('; ');
        return gap;
    }

    // ── 輔助：內容區域 ────────────────────────────────────────────────────────

    /** 取得 sheet 的內容區域（目前即 sheet 本身）。*/
    _contentArea(sheet) {
        return sheet;
    }

    /** 判斷內容區域是否溢出。+8px 吸收 sub-pixel 誤差與 LibreOffice 輸出誤差。*/
    _isOverflowing(el) {
        return el.scrollHeight > this.pageHeight + 8;
    }

    /** 判斷內容區域是否為空（無有效子節點）。*/
    _isEmpty(el) {
        const children = Array.from(el.childNodes).filter(n => {
            if (n.nodeType === Node.TEXT_NODE) return n.textContent.trim() !== '';
            if (n.nodeType === Node.ELEMENT_NODE) {
                const tag = n.tagName.toLowerCase();
                if (tag === 'br') return false;
                if (n.textContent.trim() === '' && !n.querySelector('img')) return false;
                return true;
            }
            return false;
        });
        return children.length === 0;
    }

    /** 取得內容區域的最後一個 block 節點（排除空行）。*/
    _lastBlock(el) {
        const children = Array.from(el.children);
        for (let i = children.length - 1; i >= 0; i--) {
            const child = children[i];
            // 不移動 doc-page-gap（不應出現在 content area）
            if (child.classList.contains('doc-page-gap')) continue;
            return child;
        }
        return null;
    }

    // ── Selection 保存/恢復 ───────────────────────────────────────────────────

    /**
     * 保存目前游標狀態。
     * @returns {{ startContainer, startOffset, endContainer, endOffset } | null}
     */
    _saveSelection() {
        const sel = window.getSelection();
        if (!sel || sel.rangeCount === 0) return null;
        const range = sel.getRangeAt(0);
        return {
            startContainer: range.startContainer,
            startOffset: range.startOffset,
            endContainer: range.endContainer,
            endOffset: range.endOffset,
        };
    }

    /**
     * 恢復游標至先前保存的位置。
     * 若 startContainer 已不在 DOM 中（因搬移而離開），靜默略過。
     * @param {{ startContainer, startOffset, endContainer, endOffset } | null} saved
     */
    _restoreSelection(saved) {
        if (!saved) return;
        // 確認 startContainer 仍在 container 的 DOM 樹中
        if (!this.container.contains(saved.startContainer)) return;
        try {
            const range = document.createRange();
            range.setStart(saved.startContainer, saved.startOffset);
            range.setEnd(saved.endContainer, saved.endOffset);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        } catch (_e) {
            // offset 超出範圍等邊緣情況，靜默忽略
        }
    }

    // ── Arrow Key Gap 跳過處理 ────────────────────────────────────────────────

    /**
     * 攔截 ArrowDown/Up，讓游標跳過 .doc-page-gap，避免卡住。
     * @param {KeyboardEvent} event
     */
    _skipGapOnArrow(event) {
        if (!['ArrowDown', 'ArrowUp'].includes(event.key)) return;

        const sel = window.getSelection();
        if (!sel || sel.rangeCount === 0) return;

        const range = sel.getRangeAt(0);
        const anchorEl = range.startContainer.nodeType === Node.TEXT_NODE
            ? range.startContainer.parentElement
            : range.startContainer;

        if (!anchorEl) return;

        const isDown = event.key === 'ArrowDown';

        // 找到相鄰的 .doc-page-gap
        const gap = isDown
            ? this._findNextGap(anchorEl)
            : this._findPrevGap(anchorEl);

        if (!gap) return;

        // 游標未到達頁面邊界時不攔截，讓瀏覽器正常處理
        if (!this._isCursorAtPageEdge(range, anchorEl, isDown)) return;

        event.preventDefault();

        // 目標：gap 另一側的 .doc-page-sheet 的第一/最後文字節點
        const targetSheet = isDown
            ? gap.nextElementSibling
            : gap.previousElementSibling;

        if (!targetSheet || !targetSheet.classList.contains('doc-page-sheet')) return;

        const targetNode = isDown
            ? this._firstTextNode(targetSheet)
            : this._lastTextNode(targetSheet);

        if (!targetNode) return;

        try {
            const newRange = document.createRange();
            const offset = isDown ? 0 : targetNode.length;
            newRange.setStart(targetNode, offset);
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);
        } catch (_e) {
            // 靜默忽略
        }
    }

    /**
     * 判斷游標是否在頁面的上/下邊界（用於決定是否需要跳過 gap）。
     */
    _isCursorAtPageEdge(range, anchorEl, isDown) {
        const sheet = anchorEl.closest('.doc-page-sheet');
        if (!sheet) return false;

        const sheetRect = sheet.getBoundingClientRect();
        const rangeRect = range.getBoundingClientRect();

        if (isDown) {
            // 游標在頁面下半部（距離底部不到 2 行）
            return rangeRect.bottom >= sheetRect.bottom - 40;
        } else {
            // 游標在頁面上半部（距離頂部不到 2 行）
            return rangeRect.top <= sheetRect.top + 40;
        }
    }

    _findNextGap(el) {
        let node = el;
        while (node) {
            if (node.nextElementSibling) {
                const sib = node.nextElementSibling;
                if (sib.classList && sib.classList.contains('doc-page-gap')) return sib;
                // 如果 el 在 .doc-page-sheet 內，找 sheet 的下一個兄弟
                const sheet = node.closest('.doc-page-sheet');
                if (sheet && sheet.nextElementSibling &&
                    sheet.nextElementSibling.classList.contains('doc-page-gap')) {
                    return sheet.nextElementSibling;
                }
                return null;
            }
            node = node.parentElement;
            if (node === this.container) break;
        }
        return null;
    }

    _findPrevGap(el) {
        let node = el;
        while (node) {
            const sheet = node.closest('.doc-page-sheet');
            if (sheet && sheet.previousElementSibling &&
                sheet.previousElementSibling.classList.contains('doc-page-gap')) {
                return sheet.previousElementSibling;
            }
            node = node.parentElement;
            if (node === this.container) break;
        }
        return null;
    }

    /** 取得元素內的第一個文字節點。*/
    _firstTextNode(el) {
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        const node = walker.nextNode();
        return node || null;
    }

    /** 取得元素內的最後一個文字節點。*/
    _lastTextNode(el) {
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        let last = null;
        let node;
        while ((node = walker.nextNode())) last = node;
        return last;
    }
}
