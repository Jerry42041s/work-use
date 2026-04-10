/** @odoo-module **/

import { Component, useState, useRef, useExternalListener, onMounted, onWillUnmount, onPatched } from "@odoo/owl";
import { Wysiwyg } from "@html_editor/wysiwyg";
import { MAIN_PLUGINS } from "@html_editor/plugin_sets";
import { PaginationController } from "../../core/pagination_engine";

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
 * DocPageLayout — 多頁文件容器。
 *
 * 架構說明：
 * - .doc-editor-stage（灰色背景）作為外層捲動容器
 * - 頁首/頁尾：獨立的可編輯區塊，顯示在 stage 頂端/底端
 * - 主內容：單一 Wysiwyg 編輯器，contenteditable 由 PaginationController 管理
 * - PaginationController：在 contenteditable 內建立多個 .doc-page-sheet，
 *   ResizeObserver（250ms debounce）偵測溢出並搬移 block 節點
 * - Selection/Range 在節點搬移前後自動保存/恢復（確保游標不中斷）
 * - Arrow Down/Up 自動跳過 .doc-page-gap
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
            totalPages: 1,
        });

        this._paginationCtrl = null;
        this._initAttempts = 0;

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

        onMounted(() => {
            // 使用 rAF 等待 Wysiwyg DOM 完成初始化後，再建立 PaginationController
            const tryInit = () => {
                this._initAttempts++;
                if (this._initAttempts > 50) return;  // 最多嘗試 50 次（約 800ms）

                const el = this.contentRef.el;
                if (!el) { requestAnimationFrame(tryInit); return; }

                // 找到 Wysiwyg 建立的 contenteditable
                const editable = el.querySelector('[contenteditable="true"]')
                    || el.querySelector('.odoo-editor-editable');

                if (!editable) {
                    requestAnimationFrame(tryInit);
                    return;
                }

                this._setupPaginationController(editable);
            };
            requestAnimationFrame(tryInit);
        });

        onWillUnmount(() => {
            if (this._paginationCtrl) {
                this._paginationCtrl.destroy();
                this._paginationCtrl = null;
            }
        });

        // 當邊距 props 變化時（如匯入文件後），同步更新 PaginationController 的邊距值
        // 並更新所有現有 sheet 的 CSS padding，觸發重新分頁。
        // 若 history.reset 已移除 page sheets，_redistribute 的自我修復會以新邊距重建。
        onPatched(() => {
            if (!this._paginationCtrl) return;
            const { marginTop, marginBottom, marginLeft, marginRight } = this.props;
            const ctrl = this._paginationCtrl;
            if (
                ctrl.marginTop    !== marginTop    ||
                ctrl.marginBottom !== marginBottom ||
                ctrl.marginLeft   !== marginLeft   ||
                ctrl.marginRight  !== marginRight
            ) {
                ctrl.marginTop    = marginTop;
                ctrl.marginBottom = marginBottom;
                ctrl.marginLeft   = marginLeft;
                ctrl.marginRight  = marginRight;
                ctrl.pageContentHeight = this.pageSize.height - marginTop - marginBottom;
                const padding = `${marginTop}px ${marginRight}px ${marginBottom}px ${marginLeft}px`;
                for (const sheet of ctrl._pages) {
                    sheet.style.padding = padding;
                }
                ctrl._redistributeDebounced(0);
            }
        });
    }

    // ── PaginationController 初始化 ─────────────────────────────────────────

    _setupPaginationController(editable) {
        if (this._paginationCtrl) {
            this._paginationCtrl.destroy();
        }

        const { width, height } = this.pageSize;

        this._paginationCtrl = new PaginationController({
            container: editable,
            pageWidth: width,
            pageHeight: height,
            marginTop: this.props.marginTop,
            marginBottom: this.props.marginBottom,
            marginLeft: this.props.marginLeft,
            marginRight: this.props.marginRight,
        });

        // 包裹現有內容至第一個 .doc-page-sheet，並開始分頁管理
        this._paginationCtrl.wrapExistingContent();

        // 頁數更新（PaginationController 管理）
        this._paginationCtrl.onPagesChange = (count) => {
            this.state.totalPages = count;
        };
    }

    /**
     * 當 editorConfig 或頁面格式變更時，重新初始化 PaginationController。
     * 由 DocEditor 在必要時呼叫（如匯入完成後）。
     */
    reinitPagination() {
        const el = this.contentRef.el;
        if (!el) return;
        const editable = el.querySelector('[contenteditable="true"]')
            || el.querySelector('.odoo-editor-editable');
        if (editable) this._setupPaginationController(editable);
    }

    /**
     * 取得所有頁面的序列化 HTML（去除 layout 元素）。
     * 供 DocEditor._onContentChange 使用。
     */
    getSerializedContent() {
        if (this._paginationCtrl) {
            return this._paginationCtrl.serialize();
        }
        return '';
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

    // ── 頁首/頁尾 edit mode ─────────────────────────────────────────────────

    onHeaderDblClick() {
        this.state.editingHeader = true;
    }

    onFooterDblClick() {
        this.state.editingFooter = true;
    }

    // ── 尺寸計算 ─────────────────────────────────────────────────────────────

    get pageSize() {
        return PAGE_SIZES[this.props.pageFormat] || PAGE_SIZES.A4;
    }

    get pageWidthMm() {
        const mm = { A4: 210, A3: 297, A5: 148, letter: 216, Letter: 216, legal: 216, Legal: 216 };
        return mm[this.props.pageFormat] || 210;
    }

    /** 編輯器 stage 的寬度（略大於紙張，給兩側留白）*/
    get stageStyle() {
        const { width } = this.pageSize;
        return `min-width: ${width + 128}px;`;
    }

    /** 頁首/頁尾容器（固定寬度置中，對齊紙張）*/
    get hfContainerStyle() {
        const { width } = this.pageSize;
        const { marginLeft, marginRight } = this.props;
        return [
            `width: ${width}px`,
            `margin: 0 auto`,
            `padding: 4px ${marginRight}px 4px ${marginLeft}px`,
            `box-sizing: border-box`,
        ].join('; ');
    }
}
