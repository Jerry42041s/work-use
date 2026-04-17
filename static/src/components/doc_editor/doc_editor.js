/** @odoo-module **/

import { Component, useState, onMounted, onWillUnmount } from "@odoo/owl";
import { registry } from "@web/core/registry";
import { useService } from "@web/core/utils/hooks";
import { rpc } from "@web/core/network/rpc";
import { Wysiwyg } from "@html_editor/wysiwyg";
import { MAIN_PLUGINS } from "@html_editor/plugin_sets";
import { DocPageLayout } from "../doc_page_layout/doc_page_layout";
import { DocRuler } from "../doc_ruler/doc_ruler";
import { DocOdooFieldPlugin } from "../../js/plugins/doc_odoo_field_plugin";
import { DocPageFormatPlugin } from "../../js/plugins/doc_page_format_plugin";
import { DocExportPlugin } from "../../js/plugins/doc_export_plugin";
import { DocMultiColumnPlugin } from "../../plugins/doc_multi_column_plugin";
import { DocFontFamilyPlugin } from "../../plugins/doc_font_family_plugin";
import { DocFontSizePlugin } from "../../plugins/doc_font_size_plugin";
import { DocLineHeightPlugin } from "../../plugins/doc_line_height_plugin";
import { DocTableMergePlugin } from "../../plugins/doc_table_merge_plugin";
import { DocListTypePlugin } from "../../plugins/doc_list_type_plugin";
import { AutoSaveManager } from "../../core/auto_save_manager";
import { LeaderElection } from "../../core/leader_election";
import { OfflineManager } from "../../core/offline_manager";

// 頁面寬度對照表（mm）
const PAGE_WIDTH_MM = {
    A4: 210, A3: 297, A5: 148, letter: 216, Letter: 216, legal: 216, Legal: 216,
};

export class DocEditor extends Component {
    static template = "dobtor_doc_editor.DocEditor";
    static components = { DocPageLayout, DocRuler, Wysiwyg };
    static props = ["*"];

    setup() {
        this.notification = useService("notification");
        this.action = useService("action");
        this.dialog = useService("dialog");

        // 嘗試取得 bus_service（多人協作用，可能不存在）
        try {
            this._busService = useService("bus_service");
        } catch (e) {
            this._busService = null;
        }

        this.state = useState({
            docId: null,
            docName: "未命名文件",
            editorReady: false,
            isSaving: false,
            statusMsg: "就緒",
            statusType: "saved",
            pageFormat: "A4",
            marginTop: 96,
            marginBottom: 96,
            marginLeft: 96,
            marginRight: 96,
            headerHtml: "",
            footerHtml: "",
            modelName: "",
            isOnline: true,
            showImportArea: false,
            zoom: 1.0,
            // editorConfig 放在 state 中，讓 OWL 知道變更（用於匯入強制重載）
            editorConfig: null,
        });

        this._currentContent = "";
        this._currentHeader = "";
        this._currentFooter = "";
        this.mainEditorConfig = null;
        this._editorInstance = null;   // 儲存 Wysiwyg Editor 實例，供匯入直接操作
        this._importInProgress = false; // 匯入期間抑制 onChange("") 覆蓋已設定的內容

        // 預先綁定事件處理函式，確保 addEventListener/removeEventListener 使用同一實例
        this._boundOnExportEvent = this._onExportEvent.bind(this);
        this._boundOnPageFormatEvent = this._onPageFormatEvent.bind(this);
        this._boundOnInsertPageBreak = this._onInsertPageBreak.bind(this);

        // 取得 doc_id（優先用 action context，F5 後從 sessionStorage 恢復）
        const context = this.props.action?.context || {};
        const _SESSION_KEY = "dobtor_doc_editor_last_id";
        const docId = context.doc_id || (() => {
            const stored = sessionStorage.getItem(_SESSION_KEY);
            return stored ? parseInt(stored, 10) : null;
        })();

        // ── AutoSaveManager ──
        this._autoSave = new AutoSaveManager({
            saveFn: async (html) => {
                if (!this.state.docId) return;
                await rpc("/dobtor_doc/save", {
                    doc_id: this.state.docId,
                    content_html: html,
                    header_html: this._currentHeader,
                    footer_html: this._currentFooter,
                });
            },
            debounceMs: 1500,
            maxWaitMs: 10000,
            idleMs: 3000,
            isLeaderFn: () => this._leaderElection?.isLeader() ?? true,
            onStatusChange: (status) => {
                const msgs = {
                    unsaved: ["未儲存", "saving"],
                    saving:  ["儲存中...", "saving"],
                    saved:   ["已儲存", "saved"],
                    error:   ["儲存失敗", "error"],
                };
                const [msg, type] = msgs[status] || ["就緒", "saved"];
                this.state.statusMsg = msg;
                this.state.statusType = type;
                this.state.isSaving = status === "saving";
            },
        });

        // ── OfflineManager ──
        this._offlineManager = new OfflineManager();
        this._offlineManager.onStatusChange((isOnline) => {
            this.state.isOnline = isOnline;
            if (isOnline) {
                this.notification.add("已恢復連線，正在同步...", { type: "success" });
                this._syncOfflineBuffer();
            } else {
                this.notification.add(
                    "網路已斷線，編輯內容將在恢復後自動同步",
                    { type: "warning", sticky: true }
                );
            }
        });

        // ── LeaderElection（僅在有 bus_service 且有 docId 時啟用） ──
        this._leaderElection = null;

        onMounted(async () => {
            if (docId) {
                await this._loadDocument(docId);
            } else {
                this._initEditorConfig("");
                this.state.editorReady = true;
            }

            // 初始化 Leader Election
            if (this._busService && this.state.docId) {
                const channel = `doc.document_${this.state.docId}`;
                const sessionId = Math.random().toString(36).slice(2);
                this._leaderElection = new LeaderElection(this._busService, channel, sessionId);
            }

            document.addEventListener("doc-export", this._boundOnExportEvent);
            document.addEventListener("doc-page-format-change", this._boundOnPageFormatEvent);
            document.addEventListener("doc-insert-page-break", this._boundOnInsertPageBreak);
        });

        onWillUnmount(async () => {
            await this._autoSave.flush();
            this._autoSave.destroy();
            this._offlineManager.destroy();
            if (this._leaderElection) this._leaderElection.destroy();

            document.removeEventListener("doc-export", this._boundOnExportEvent);
            document.removeEventListener("doc-page-format-change", this._boundOnPageFormatEvent);
            document.removeEventListener("doc-insert-page-break", this._boundOnInsertPageBreak);
        });
    }

    // ─── 資料載入 ────────────────────────────────────────────────

    async _loadDocument(docId) {
        try {
            const data = await rpc("/dobtor_doc/load", { doc_id: docId });
            this.state.docId = data.id;
            this.state.docName = data.name;
            // 記錄目前文件 ID，讓 F5 後可恢復
            sessionStorage.setItem("dobtor_doc_editor_last_id", data.id);
            this.state.pageFormat = data.page_format || "A4";
            this.state.marginTop = data.margin_top || 96;
            this.state.marginBottom = data.margin_bottom || 96;
            this.state.marginLeft = data.margin_left || 96;
            this.state.marginRight = data.margin_right || 96;
            this.state.headerHtml = data.header_html || "";
            this.state.footerHtml = data.footer_html || "";
            this.state.modelName = data.model_name || "";

            this._currentContent = data.content_html || "";
            this._currentHeader = data.header_html || "";
            this._currentFooter = data.footer_html || "";

            this._initEditorConfig(this._currentContent);
            this.state.editorReady = true;
            this.state.statusMsg = "已載入";
            this.state.statusType = "saved";
        } catch (error) {
            this.state.statusMsg = `載入失敗：${error.message || error}`;
            this.state.statusType = "error";
            console.error("[DocEditor] Load failed:", error);
        }
    }

    _initEditorConfig(initialContent) {
        const config = {
            Plugins: [
                ...MAIN_PLUGINS,
                DocOdooFieldPlugin,
                DocPageFormatPlugin,
                DocExportPlugin,
                DocMultiColumnPlugin,
                DocFontFamilyPlugin,
                DocFontSizePlugin,
                DocLineHeightPlugin,
                DocTableMergePlugin,
                DocListTypePlugin,
            ],
            content: initialContent || "",
            onChange: (html) => this._onContentChange(html),
            // onLoad：編輯器初始化完成後取得實例，供後續 history.reset() 使用
            onLoad: (editor) => {
                this._editorInstance = editor;
            },
            placeholder: "開始輸入...",
            dobtor_doc_id: this.state.docId,
            dobtor_model_name: this.state.modelName,
        };
        // 同時更新 reactive state（供匯入時 OWL 感知）與非 reactive 參考
        this.mainEditorConfig = config;
        this.state.editorConfig = config;
    }

    // ─── 內容變更 ─────────────────────────────────────────────────

    /**
     * 移除 layout 元素，回傳乾淨 HTML 以存入 DB：
     * - .doc-page-separator（舊版視覺分頁線）
     * - .doc-page-gap（新版頁間間隔，non-editable）
     * - .doc-page-sheet（新版頁面容器，保留其內容）
     */
    _stripLayoutElements(html) {
        if (!html) return html;
        const hasLayout = html.includes('doc-page-separator') ||
                          html.includes('doc-page-gap') ||
                          html.includes('doc-page-sheet');
        if (!hasLayout) return html;

        const tmp = document.createElement('div');
        tmp.innerHTML = html;

        // 移除 .doc-page-separator 與 .doc-page-gap
        tmp.querySelectorAll('.doc-page-separator, .doc-page-gap').forEach(el => el.remove());

        // 拆包 .doc-page-sheet（保留子節點，移除外層 div）
        tmp.querySelectorAll('.doc-page-sheet').forEach(sheet => {
            const parent = sheet.parentNode;
            if (!parent) return;
            while (sheet.firstChild) {
                parent.insertBefore(sheet.firstChild, sheet);
            }
            parent.removeChild(sheet);
        });

        return tmp.innerHTML;
    }

    // 保留舊名稱作為別名（向後相容，避免其他地方呼叫出錯）
    _stripPageSeparators(html) {
        return this._stripLayoutElements(html);
    }

    _onContentChange(html) {
        // 匯入期間：忽略空白的 onChange（Wysiwyg unmount 時可能觸發）
        if (this._importInProgress && !html.trim()) {
            return;
        }
        // 移除 layout 元素（不儲存至 DB）
        this._currentContent = this._stripLayoutElements(html);
        if (this._offlineManager.isOnline) {
            this._autoSave.onContentChange(html);
        } else {
            this._offlineManager.bufferOperation({ type: "save", html });
            this.state.statusMsg = "離線緩存中";
            this.state.statusType = "saving";
        }
    }

    onHeaderChange(html) {
        this._currentHeader = html;
        this._autoSave.onContentChange(this._currentContent);
    }

    onFooterChange(html) {
        this._currentFooter = html;
        this._autoSave.onContentChange(this._currentContent);
    }

    async _syncOfflineBuffer() {
        const ops = this._offlineManager.drainBuffer();
        if (!ops.length || !this.state.docId) return;
        const lastSave = [...ops].reverse().find(op => op.type === "save");
        if (!lastSave) return;
        try {
            await rpc("/dobtor_doc/save", {
                doc_id: this.state.docId,
                content_html: lastSave.html,
                header_html: this._currentHeader,
                footer_html: this._currentFooter,
            });
            this.state.statusMsg = "已同步";
            this.state.statusType = "saved";
        } catch (e) {
            this.notification.add(`同步失敗：${e.message}`, { type: "danger" });
        }
    }

    // ─── 手動儲存 ────────────────────────────────────────────────

    async onSave() {
        if (!this.state.docId || this.state.isSaving) return;
        sessionStorage.setItem("dobtor_doc_editor_last_id", this.state.docId);
        this.state.isSaving = true;
        this.state.statusMsg = "儲存中...";
        this.state.statusType = "saving";
        try {
            await rpc("/dobtor_doc/save", {
                doc_id: this.state.docId,
                content_html: this._currentContent,
                header_html: this._currentHeader,
                footer_html: this._currentFooter,
            });
            this.state.statusMsg = "已儲存";
            this.state.statusType = "saved";
        } catch (error) {
            this.state.statusMsg = `儲存失敗：${error.message || error}`;
            this.state.statusType = "error";
            this.notification.add("文件儲存失敗", { type: "danger" });
        } finally {
            this.state.isSaving = false;
        }
    }

    // ─── 匯出 ────────────────────────────────────────────────────

    async onExport(format, quality = "high") {
        if (!this.state.docId) {
            this.notification.add("請先儲存文件", { type: "warning" });
            return;
        }
        await this.onSave();
        this.state.statusMsg = `正在產生 ${format.toUpperCase()}...`;
        this.state.statusType = "saving";
        try {
            const result = await rpc("/dobtor_doc/export", {
                doc_id: this.state.docId,
                format: format,
                quality: quality,
            });
            if (result.error) throw new Error(result.error);
            const blob = this._base64ToBlob(result.data, result.mimetype);
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = result.filename;
            a.click();
            URL.revokeObjectURL(url);
            this.state.statusMsg = "已儲存";
            this.state.statusType = "saved";
            this.notification.add(`${result.filename} 下載中`, { type: "success" });
        } catch (error) {
            this.state.statusMsg = "匯出失敗";
            this.state.statusType = "error";
            this.notification.add(`匯出失敗：${error.message || error}`, { type: "danger" });
        }
    }

    // ─── 匯入 ────────────────────────────────────────────────────

    onImportClick() {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = ".docx,.odt";
        input.onchange = (ev) => this._handleImportFile(ev.target.files[0]);
        input.click();
    }

    async _handleImportFile(file) {
        if (!file) return;
        this.state.statusMsg = "匯入中...";
        this.state.statusType = "saving";
        try {
            const formData = new FormData();
            formData.append("file", file);
            const response = await fetch("/dobtor_doc/import", {
                method: "POST",
                body: formData,
            });
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            const data = await response.json();
            if (data.error) {
                this.notification.add(`匯入失敗：${data.error}`, { type: "danger" });
            } else if (data.html !== undefined) {
                if (!data.html.trim()) {
                    this.notification.add("匯入完成，但文件內容為空", { type: "warning" });
                    return;
                }

                // ── 立即更新 _currentContent，確保後續手動儲存使用正確內容 ──
                this._currentContent = data.html;

                // ── 從檔名擷取文件名稱（去除副檔名）──
                const importedName = file.name.replace(/\.[^/.]+$/, "") || "匯入文件";
                this.state.docName = importedName;

                // ── 套用原始文件的頁面邊距（若 LibreOffice 有解析到）──
                if (data.margins && this.state.docId) {
                    const m = data.margins;
                    if (m.top)    this.state.marginTop    = m.top;
                    if (m.bottom) this.state.marginBottom = m.bottom;
                    if (m.left)   this.state.marginLeft   = m.left;
                    if (m.right)  this.state.marginRight  = m.right;
                    rpc("/dobtor_doc/save_settings", {
                        doc_id:        this.state.docId,
                        margin_top:    m.top    || this.state.marginTop,
                        margin_bottom: m.bottom || this.state.marginBottom,
                        margin_left:   m.left   || this.state.marginLeft,
                        margin_right:  m.right  || this.state.marginRight,
                    }).catch(e => console.warn('[DocEditor] margin save failed:', e));
                }

                // ── 取消匯入前的舊存檔計時器，防止舊內容覆蓋匯入內容 ──
                this._autoSave.cancel();

                // ── 替換 editor 內容 ──
                if (this._editorInstance?.shared?.history) {
                    // 標準路徑：透過 history.reset 更新編輯器內容
                    this._editorInstance.shared.history.reset(data.html);
                    this._onContentChange(data.html);
                } else {
                    // Fallback：editor 尚未 ready，unmount 後以新內容重新掛載
                    // 設定旗標，抑制 Wysiwyg unmount 時可能觸發的 onChange("") 覆蓋
                    this._importInProgress = true;
                    this.state.editorConfig = null;
                    await new Promise(r => requestAnimationFrame(r));
                    await new Promise(r => setTimeout(r, 0));
                    this._initEditorConfig(data.html);
                }

                // ── 立即儲存到 DB（確保 F5 後仍能看到匯入內容）──
                if (this.state.docId) {
                    try {
                        await rpc("/dobtor_doc/save", {
                            doc_id: this.state.docId,
                            content_html: data.html,
                            header_html: this._currentHeader,
                            footer_html: this._currentFooter,
                            name: importedName,
                        });
                        // 告知 AutoSave 目前已儲存的內容，防止稍後的自動存檔覆蓋
                        this._autoSave._lastSavedHtml = data.html;
                    } catch (saveErr) {
                        console.warn("[DocEditor] import save failed:", saveErr);
                    }
                }

                // ── 解除匯入旗標（允許後續正常的 onChange 更新） ──
                this._importInProgress = false;

                this.state.statusMsg = "已儲存";
                this.state.statusType = "saved";
                this.notification.add("匯入成功，正在重新整理畫面...", { type: "success" });

                // 重新載入頁面確保 Pagination 引擎乾淨掛載，避免 DOM 狀態不同步
                setTimeout(() => {
                    window.location.reload();
                }, 500);
            } else {
                this.state.statusMsg = "就緒";
                this.state.statusType = "saved";
                this.notification.add("匯入失敗：無法解析檔案內容", { type: "danger" });
            }
        } catch (e) {
            this.state.statusMsg = "就緒";
            this.state.statusType = "saved";
            this.notification.add(`匯入失敗：${e.message}`, { type: "danger" });
        }
    }

    _base64ToBlob(base64, mimeType) {
        const bytes = atob(base64);
        const buf = new Uint8Array(bytes.length);
        for (let i = 0; i < bytes.length; i++) buf[i] = bytes.charCodeAt(i);
        return new Blob([buf], { type: mimeType });
    }

    // ─── Plugin 事件處理 ──────────────────────────────────────────

    _onExportEvent(event) {
        const { action } = event.detail || {};
        if (action === "export-pdf") this.onExport("pdf");
        else if (action === "export-docx-hq") this.onExport("docx", "high");
        else if (action === "import-docx") this.onImportClick();
    }

    _onPageFormatEvent(event) {
        const { format } = event.detail || {};
        if (format && PAGE_WIDTH_MM[format] !== undefined) {
            this.state.pageFormat = format;
            if (this.state.docId) {
                rpc("/dobtor_doc/save_settings", {
                    doc_id: this.state.docId,
                    page_format: format,
                }).catch(() => {});
            }
        }
    }

    _onInsertPageBreak() {
        this._autoSave.onContentChange(this._currentContent);
    }

    // ─── Toolbar 事件 ─────────────────────────────────────────────

    onTitleChange(event) {
        const newName = event.target.value.trim() || "未命名文件";
        this.state.docName = newName;
        if (this.state.docId) {
            rpc("/dobtor_doc/save", {
                doc_id: this.state.docId,
                name: newName,
            }).catch(() => {});
        }
    }

    onPageFormatChange(event) {
        const format = event.target.value;
        this._onPageFormatEvent({ detail: { format } });
    }

    onClose() {
        history.back();
    }

    onZoomChange(event) {
        this.state.zoom = parseFloat(event.target.value) || 1.0;
    }

    // ─── 版本歷史 ─────────────────────────────────────────────────

    async onSaveVersion() {
        if (!this.state.docId) return;
        await this.onSave();
        try {
            await rpc("/dobtor_doc/save_version", { doc_id: this.state.docId });
            this.notification.add("版本已儲存", { type: "success" });
        } catch (e) {
            this.notification.add(`版本儲存失敗：${e.message}`, { type: "danger" });
        }
    }

    // ─── 工具方法 ────────────────────────────────────────────────

    get statusClass() {
        const map = {
            saved:  "doc-statusbar-saved",
            saving: "doc-statusbar-saving",
            error:  "doc-statusbar-error",
        };
        return map[this.state.statusType] || "";
    }

    get pageWidthMm() {
        return PAGE_WIDTH_MM[this.state.pageFormat] || 210;
    }

    get offlineBadge() {
        return !this.state.isOnline;
    }
}

registry.category("actions").add("dobtor_doc_editor.action_doc_editor", DocEditor);
