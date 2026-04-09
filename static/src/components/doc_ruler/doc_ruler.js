/** @odoo-module **/

import { Component } from "@odoo/owl";

// A4 寬度 210mm，96dpi 下 1mm ≈ 3.7795px
const MM_TO_PX = 96 / 25.4;

/**
 * DocRuler — 純 CSS 水平刻度尺（V0.1，無拖曳互動）。
 * V0.2 計畫替換為 @scena/ruler 實現拖曳設定邊距。
 */
export class DocRuler extends Component {
    static template = "dobtor_doc_editor.DocRuler";
    static props = {
        pageWidthMm: { type: Number, default: 210 },  // 頁面寬度（mm）
        unit: { type: String, default: "mm" },
    };

    get paperWidthPx() {
        return Math.round(this.props.pageWidthMm * MM_TO_PX);
    }

    /**
     * 產生刻度清單。
     * 每 5mm 一小刻，每 10mm 一大刻，每 20mm 顯示標籤。
     */
    get ticks() {
        const totalMm = this.props.pageWidthMm;
        const ticks = [];
        for (let mm = 0; mm <= totalMm; mm += 5) {
            ticks.push({
                pos: mm,
                px: Math.round(mm * MM_TO_PX),
                major: mm % 10 === 0,
                label: mm % 20 === 0 ? mm : null,
            });
        }
        return ticks;
    }
}
