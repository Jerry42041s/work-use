/** @odoo-module **/

import { Component, useState, useRef, useExternalListener, onMounted, onWillUnmount } from "@odoo/owl";
import { Wysiwyg } from "@html_editor/wysiwyg";
import { MAIN_PLUGINS } from "@html_editor/plugin_sets";
import { PaginationEngine } from "../../core/pagination_engine";

// 頁面尺寸 @ 96 dpi（px）
const PAGE_SIZES = {
    A4:     { width: 794,  height: 1123 },
    A3:     { width: 1123, height: 1587 },
    A5:     { width: 559,  height: 794  },
    letter: { width: 816,  height: 1056 },
    Letter: { width: 816,  height: 1056 },
    legal:  { width: 816,  height: 1344 },
    Legal:  { width: 816,  height: 1344 },
};

/**
 * DocPageLayout — 頁面容器。
 *
 * Google Docs 正確模型：
 * - 頁首/頁尾在「邊距區域」內（絕對定位），不佔用內容空間
 * - 預設 readonly（顯示 innerHTML）
 * - 雙擊進入 edit mode（掛載 Wysiwyg）
 * - 點擊外部退出 edit mode
 *
 * 結構：
 * .doc-page-sheet (position: relative; padding: marginTop...marginBottom)
 *   .doc-header-area  (position: absolute; top: 0; height: marginTop)
 *   .doc-content-wrapper (normal flow, fills full padding area)
 *   .doc-footer-area  (position: absolute; bottom: 0; height: marginBottom)
 */
export class DocPageLayout extends Component {
    static template = "dobtor_doc_editor.DocPageLayout";
    static components = { Wysiwyg };
    static props = {
        pageFormat:     { type: String, default: "A4" },
        marginTop:      { type: Number, default: 96 },
        marginBottom:   { type: Number, default: 96 },
        marginLeft:     { type: Number, default: 96 },
        marginRight:    { type: Number, default: 96 },
        headerHtml:     { type: String, optional: true },
        footerHtml:     { type: String, optional: true },
        editorConfig:   { type: Object, optional: true },
        onHeaderChange: { type: Function, optional: true },
        onFooterChange: { type: Function, optional: true },
    };

    setup() {
        this.contentRef  = useRef("pageContent");
        this.headerRef   = useRef("headerArea");
        this.footerRef   = useRef("footerArea");

        this.state = useState({
            editingHeader: false,
            editingFooter: false,
            pageBreaks:    [],   // [{ pageIndex, topPx }]
            totalPages:    1,
        });

        // 頁首/頁尾 Wysiwyg 設定由 getter 動態產生，確保每次掛載時取得最新 props.content

        // 點擊頁首/頁尾外部時退出 edit mode
        useExternalListener(document, "mousedown", (ev) => {
            if (this.state.editingHeader) {
                const el = this.headerRef.el;
                if (el && !el.contains(ev.target)) {
                    this.state.editingHeader = false;
                }
            }
            if (this.state.editingFooter) {
                const el = this.footerRef.el;
                if (el && !el.contains(ev.target)) {
                    this.state.editingFooter = false;
                }
            }
        });

        this._resizeObserver = null;
        this._paginationEngine = null;
        this._isRecalculating = false;

        onMounted(() => {
            this._resizeObserver = new ResizeObserver(() => this._recalcPages());
            if (this.contentRef.el) {
                this._resizeObserver.observe(this.contentRef.el);
            }
        });

        onWillUnmount(() => {
            if (this._resizeObserver) this._resizeObserver.disconnect();
        });
    }

    // ── 頁首/頁尾 Wysiwyg 設定（getter 確保每次掛載時取用最新 props）──

    get headerEditorConfig() {
        return {
            Plugins: MAIN_PLUGINS,
            content: this.props.headerHtml || "",
            onChange: (html) => {
                if (this.props.onHeaderChange) this.props.onHeaderChange(html);
            },
            placeholder: "頁首...",
        };
    }

    get footerEditorConfig() {
        return {
            Plugins: MAIN_PLUGINS,
            content: this.props.footerHtml || "",
            onChange: (html) => {
                if (this.props.onFooterChange) this.props.onFooterChange(html);
            },
            placeholder: "頁尾...",
        };
    }

    // ── 頁首/頁尾 edit mode ─────────────────────────────────────

    onHeaderDblClick() {
        this.state.editingHeader = true;
    }

    onFooterDblClick() {
        this.state.editingFooter = true;
    }

    // ── 尺寸計算 ────────────────────────────────────────────────

    get pageSize() {
        return PAGE_SIZES[this.props.pageFormat] || PAGE_SIZES.A4;
    }

    get pageWidthMm() {
        const mm = { A4: 210, A3: 297, A5: 148, letter: 216, Letter: 216, legal: 216, Legal: 216 };
        return mm[this.props.pageFormat] || 210;
    }

    /** 紙張容器：relative + padding 形成邊距 + 固定最小頁面高度
     *  不使用 flex：讓 .doc-content-wrapper 自然隨內容增長，頁面可超過一頁高度 */
    get pageStyle() {
        const { marginTop, marginRight, marginBottom, marginLeft } = this.props;
        const { height } = this.pageSize;
        return [
            `position: relative`,
            `padding: ${marginTop}px ${marginRight}px ${marginBottom}px ${marginLeft}px`,
            `min-height: ${height}px`,
            `box-sizing: border-box`,
        ].join("; ");
    }

    /** 頁首絕對定位至上邊距區域 */
    get headerStyle() {
        const { marginTop, marginLeft, marginRight } = this.props;
        return [
            `position: absolute`,
            `top: 0`,
            `left: 0`,
            `right: 0`,
            `height: ${marginTop}px`,
            `padding: 4px ${marginRight}px 4px ${marginLeft}px`,
            `box-sizing: border-box`,
        ].join("; ");
    }

    /** 頁尾絕對定位至下邊距區域 */
    get footerStyle() {
        const { marginBottom, marginLeft, marginRight } = this.props;
        return [
            `position: absolute`,
            `bottom: 0`,
            `left: 0`,
            `right: 0`,
            `height: ${marginBottom}px`,
            `padding: 4px ${marginRight}px 4px ${marginLeft}px`,
            `box-sizing: border-box`,
        ].join("; ");
    }

    /** 使用 PaginationEngine 計算分頁，並以 DOM spacer 隔開各頁（避免 absolute 元素遮蓋內容） */
    _recalcPages() {
        if (this._isRecalculating) return;
        this._isRecalculating = true;
        try {
            const el = this.contentRef.el;
            if (!el) return;

            const { height } = this.pageSize;
            const pageContentH = height - this.props.marginTop - this.props.marginBottom;
            if (pageContentH <= 0) return;

            const editable = el.querySelector('[contenteditable="true"]')
                || el.querySelector('.odoo-editor-editable')
                || el;

            // 1. 移除現有 separator（讓 PaginationEngine 看到純淨內容）
            Array.from(editable.querySelectorAll('.doc-page-separator'))
                .forEach(s => s.remove());

            // 2. 惰性建立或重建 PaginationEngine
            if (!this._paginationEngine ||
                this._paginationEngine.pageContentHeight !== pageContentH) {
                this._paginationEngine = new PaginationEngine({ pageContentHeight: pageContentH });
            }

            // 3. 計算分頁（PaginationEngine 會跳過 .doc-page-separator）
            const pages = this._paginationEngine.computePageBreaks(editable);

            if (pages.length <= 1) {
                this.state.pageBreaks = [];
                this.state.totalPages = 1;
                return;
            }

            // 4. 在各頁起始 block 前插入視覺間隔 div
            for (const p of pages.slice(1)) {
                const sep = this._createPageSeparator(p.pageIndex);
                if (p.blockRef && p.blockRef.parentNode === editable) {
                    editable.insertBefore(sep, p.blockRef);
                }
            }

            this.state.totalPages = pages.length;
            this.state.pageBreaks = []; // DOM spacer 取代 absolute break visual
        } finally {
            this._isRecalculating = false;
        }
    }

    /** 建立頁面間隔元素（contenteditable=false，不存入 DB） */
    _createPageSeparator(pageIndex) {
        const ml = this.props.marginLeft;
        const mr = this.props.marginRight;
        const sep = document.createElement('div');
        sep.contentEditable = 'false';
        sep.className = 'doc-page-separator';
        sep.setAttribute('data-page-separator', 'true');
        sep.style.cssText =
            `height: 32px; ` +
            `margin-left: -${ml}px; margin-right: -${mr}px; ` +
            `background: #e8e8e8; ` +
            `box-shadow: inset 0 4px 6px -4px rgba(0,0,0,0.15), ` +
            `inset 0 -4px 6px -4px rgba(0,0,0,0.15); ` +
            `user-select: none; pointer-events: none; position: relative;`;
        sep.innerHTML =
            `<span style="position:absolute;top:50%;left:50%;` +
            `transform:translate(-50%,-50%);font-size:0.7rem;color:#999;` +
            `white-space:nowrap;user-select:none;">第 ${pageIndex} 頁</span>`;
        return sep;
    }
}
