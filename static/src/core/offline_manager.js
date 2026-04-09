/** @odoo-module **/

/**
 * OfflineManager — 監聽網路狀態，離線時 buffer 操作，上線後通知。
 */
export class OfflineManager {
    constructor() {
        this._online = navigator.onLine;
        this._buffer = [];
        this._callbacks = [];
        this._onOnline = this._handleOnline.bind(this);
        this._onOffline = this._handleOffline.bind(this);

        window.addEventListener("online", this._onOnline);
        window.addEventListener("offline", this._onOffline);
    }

    get isOnline() {
        return this._online;
    }

    onStatusChange(callback) {
        this._callbacks.push(callback);
    }

    bufferOperation(op) {
        this._buffer.push(op);
    }

    drainBuffer() {
        const ops = [...this._buffer];
        this._buffer = [];
        return ops;
    }

    _handleOnline() {
        this._online = true;
        for (const cb of this._callbacks) cb(true);
    }

    _handleOffline() {
        this._online = false;
        for (const cb of this._callbacks) cb(false);
    }

    destroy() {
        window.removeEventListener("online", this._onOnline);
        window.removeEventListener("offline", this._onOffline);
    }
}
