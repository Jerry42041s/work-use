/** @odoo-module **/

import { Plugin } from "@html_editor/plugin";
import { _t } from "@web/core/l10n/translation";
import { withSequence } from "@html_editor/utils/resource";
import { rpc } from "@web/core/network/rpc";
import { getHTMLtoDOCX } from "../../core/lazy_loader";

/**
 * DocExportPlugin — 提供匯出指令與 DOCX 匯入功能。
 * 匯出分兩個品質等級：草稿（前端）/ 高品質（後端 LibreOffice）。
 */
export class DocExportPlugin extends Plugin {
    static id = "docExport";
    static dependencies = [];

    get resources() {
        return {
            powerbox_categories: [
                withSequence(90, {
                    id: "docExport",
                    name: _t("匯出 / 匯入"),
                }),
            ],
            user_commands: [
                {
                    id: "exportPdf",
                    title: _t("匯出 PDF"),
                    description: _t("高品質 PDF（含頁首頁尾）"),
                    icon: "fa-file-pdf-o",
                    run: () => this._triggerEvent("export-pdf"),
                },
                {
                    id: "exportDocxHQ",
                    title: _t("高品質匯出 DOCX"),
                    description: _t("含頁首頁尾，需數秒"),
                    icon: "fa-file-word-o",
                    run: () => this._triggerEvent("export-docx-hq"),
                },
                {
                    id: "exportDocxDraft",
                    title: _t("快速匯出 DOCX（草稿）"),
                    description: _t("前端即時，適合快速分享"),
                    icon: "fa-file-word-o",
                    run: () => this._exportQuickDocx(),
                },
                {
                    id: "importDocx",
                    title: _t("匯入 DOCX"),
                    description: _t("從 Word 檔案匯入內容"),
                    icon: "fa-upload",
                    run: () => this._triggerEvent("import-docx"),
                },
            ],
            powerbox_items: [
                {
                    categoryId: "docExport",
                    commandId: "exportPdf",
                    title: _t("匯出 PDF"),
                    description: _t("最終成品"),
                    icon: "fa-file-pdf-o",
                    keywords: ["pdf", "export", "匯出"],
                },
                {
                    categoryId: "docExport",
                    commandId: "exportDocxHQ",
                    title: _t("高品質 DOCX"),
                    description: _t("含頁首頁尾"),
                    icon: "fa-file-word-o",
                    keywords: ["docx", "word", "高品質"],
                },
                {
                    categoryId: "docExport",
                    commandId: "exportDocxDraft",
                    title: _t("快速 DOCX"),
                    description: _t("草稿級即時下載"),
                    icon: "fa-file-word-o",
                    keywords: ["docx", "draft", "草稿"],
                },
                {
                    categoryId: "docExport",
                    commandId: "importDocx",
                    title: _t("匯入 DOCX"),
                    description: _t("從 Word 匯入"),
                    icon: "fa-upload",
                    keywords: ["import", "匯入", "docx"],
                },
            ],
        };
    }

    _triggerEvent(action) {
        if (this.editable) {
            this.editable.dispatchEvent(new CustomEvent("doc-export", {
                detail: { action },
                bubbles: true,
                cancelable: false,
            }));
        }
    }

    async _exportQuickDocx() {
        const notification = this.services?.notification;
        try {
            if (notification) notification.add(_t("正在載入匯出元件..."), { type: "info" });
            const HTMLtoDOCX = await getHTMLtoDOCX();
            const html = this.editable?.innerHTML || "";
            const blob = await HTMLtoDOCX(html, null, {
                orientation: "portrait",
                margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
                font: "Microsoft JhengHei",
                fontSize: 24,
            });
            this._downloadBlob(blob, "document_draft.docx");
            if (notification) notification.add(_t("草稿 DOCX 下載中"), { type: "success" });
        } catch (err) {
            if (notification) {
                notification.add(
                    `${_t("快速 DOCX 匯出失敗：")}${err.message}`,
                    { type: "danger" }
                );
            }
        }
    }

    _downloadBlob(blob, filename) {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }
}
