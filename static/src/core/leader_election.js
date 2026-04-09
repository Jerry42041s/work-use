/** @odoo-module **/

/**
 * LeaderElection — 基於 bus.bus presence 的簡單 Leader Election。
 * 規則：最早加入的 session 為 leader。
 * 當 leader 離開時，次早的 session 自動接任。
 *
 * 用於多人協作時防止 race condition：只有 leader 負責 autosave。
 */
export class LeaderElection {
    constructor(busService, documentChannel, sessionId) {
        this.busService = busService;
        this.documentChannel = documentChannel;
        this.sessionId = sessionId;
        this._presences = new Map();
        this._joinTimestamp = Date.now();
        this._bound = this._onNotification.bind(this);

        // 記錄自己的 presence
        this._presences.set(sessionId, this._joinTimestamp);

        if (this.busService) {
            this.busService.addEventListener("notification", this._bound);
            this._announce();
        }
    }

    _announce() {
        try {
            this.busService.send(this.documentChannel, {
                type: "doc_editor_presence",
                sessionId: this.sessionId,
                joinTimestamp: this._joinTimestamp,
            });
        } catch (e) {
            // bus.bus 未啟用時靜默失敗
        }
    }

    _onNotification({ detail: notifications }) {
        if (!Array.isArray(notifications)) return;
        for (const { payload } of notifications) {
            if (!payload) continue;
            if (payload.type === "doc_editor_presence") {
                this._presences.set(payload.sessionId, payload.joinTimestamp);
            }
            if (payload.type === "doc_editor_leave") {
                this._presences.delete(payload.sessionId);
            }
        }
    }

    isLeader() {
        if (this._presences.size <= 1) return true;
        const sorted = [...this._presences.entries()].sort((a, b) => a[1] - b[1]);
        return sorted[0][0] === this.sessionId;
    }

    destroy() {
        if (this.busService) {
            this.busService.removeEventListener("notification", this._bound);
            try {
                this.busService.send(this.documentChannel, {
                    type: "doc_editor_leave",
                    sessionId: this.sessionId,
                });
            } catch (e) {}
        }
    }
}
