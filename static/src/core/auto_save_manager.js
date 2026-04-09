/** @odoo-module **/

/**
 * AutoSaveManager — 三層保護機制：
 * Layer 1: Debounce（快速輸入不頻繁觸發）
 * Layer 2: Max Wait（持續輸入超時強制存）
 * Layer 3: Idle Detection（停止輸入後延遲存檔）
 *
 * 搭配 LeaderElection 使用時，多人協作只有 leader 負責存檔。
 */
export class AutoSaveManager {
    constructor(options) {
        this.saveFn = options.saveFn;
        this.debounceMs = options.debounceMs ?? 1500;
        this.maxWaitMs = options.maxWaitMs ?? 10000;
        this.idleMs = options.idleMs ?? 3000;
        this.isLeaderFn = options.isLeaderFn ?? (() => true);
        this.onStatusChange = options.onStatusChange ?? (() => {});

        this._debounceTimer = null;
        this._idleTimer = null;
        this._maxWaitTimer = null;
        this._firstChangeTime = null;
        this._saving = false;
        this._pendingSave = null;
        this._lastSavedHtml = null;
        this._destroyed = false;
    }

    onContentChange(html) {
        if (this._destroyed) return;
        if (html === this._lastSavedHtml) return;

        this.onStatusChange("unsaved");

        // Layer 1: Debounce
        clearTimeout(this._debounceTimer);
        this._debounceTimer = setTimeout(() => this._tryTriggerSave(html), this.debounceMs);

        // Layer 2: Max Wait
        if (!this._firstChangeTime) {
            this._firstChangeTime = Date.now();
            this._maxWaitTimer = setTimeout(() => {
                this._tryTriggerSave(html);
                this._firstChangeTime = null;
            }, this.maxWaitMs);
        }

        // Layer 3: Idle Detection
        clearTimeout(this._idleTimer);
        this._idleTimer = setTimeout(() => this._tryTriggerSave(html), this.idleMs);
    }

    async _tryTriggerSave(html) {
        if (this._destroyed) return;
        clearTimeout(this._debounceTimer);
        clearTimeout(this._idleTimer);
        clearTimeout(this._maxWaitTimer);
        this._firstChangeTime = null;

        if (!this.isLeaderFn()) return;

        if (this._saving) {
            this._pendingSave = html;
            return;
        }

        this._saving = true;
        this.onStatusChange("saving");
        try {
            await this.saveFn(html);
            this._lastSavedHtml = html;
            this.onStatusChange("saved");
        } catch (err) {
            console.error("[AutoSave] Save failed:", err);
            this.onStatusChange("error");
            // 3 秒後重試
            if (!this._destroyed) {
                setTimeout(() => this.onContentChange(html), 3000);
            }
        } finally {
            this._saving = false;
            if (this._pendingSave !== null) {
                const pending = this._pendingSave;
                this._pendingSave = null;
                this.onContentChange(pending);
            }
        }
    }

    /** 強制立即儲存（元件卸載前呼叫） */
    async flush() {
        if (this._destroyed || !this._pendingSave) return;
        clearTimeout(this._debounceTimer);
        clearTimeout(this._idleTimer);
        clearTimeout(this._maxWaitTimer);
        const pending = this._pendingSave;
        this._pendingSave = null;
        if (pending && this.isLeaderFn()) {
            await this.saveFn(pending).catch(() => {});
        }
    }

    /** 取消所有待執行的計時器（匯入時呼叫，清除舊內容的存檔任務） */
    cancel() {
        clearTimeout(this._debounceTimer);
        clearTimeout(this._idleTimer);
        clearTimeout(this._maxWaitTimer);
        this._firstChangeTime = null;
        this._debounceTimer = null;
        this._idleTimer = null;
        this._maxWaitTimer = null;
    }

    destroy() {
        this._destroyed = true;
        clearTimeout(this._debounceTimer);
        clearTimeout(this._idleTimer);
        clearTimeout(this._maxWaitTimer);
    }
}
