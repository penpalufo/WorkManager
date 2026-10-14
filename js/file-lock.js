// NAS上の排他ロックを取得し、編集中は定期的に有効期限を更新する
export const fileLockMethods = {
	workManagerSaveUrl() {
		// メイン画面の保守用指定を、保存APIへ明示的に引き継ぐ。
		return './php/save.php' + (this.ignoreFileLock ? '?ignoreLock=1' : '');
	},

	async acquireFileLock() {
		// 起動時に編集権を取得する。取得できなくても一覧の閲覧は可能。
		if (this.ignoreFileLock) {
			this.hasFileLock = true;
			this.isFileLocked = false;
			this.lockMessage = '';
			return;
		}
		try {
			const response = await fetch('./php/lock.php?action=acquire&client=' + encodeURIComponent(this.lockClientId), { method: 'POST' });
			const result = await this.getJsonResponse(response);
			this.hasFileLock = response.ok && result.success;
			this.isFileLocked = !this.hasFileLock;
			this.lockMessage = this.hasFileLock ? '' : (result.message || 'ファイルロック中です。編集はできません。');
		} catch (error) {
			this.hasFileLock = false;
			this.isFileLocked = true;
			this.lockMessage = 'ファイルロックを確認できないため、編集はできません。';
		}
	},

	async refreshFileLock() {
		// 長時間編集中でもロックが期限切れにならないよう定期更新する。
		if (this.ignoreFileLock) return;
		if (!this.hasFileLock) return;
		try {
			const response = await fetch('./php/lock.php?action=refresh&client=' + encodeURIComponent(this.lockClientId), { method: 'POST' });
			const result = await this.getJsonResponse(response);
			if (!response.ok || !result.success) {
				this.hasFileLock = false;
				this.isFileLocked = true;
				this.lockMessage = result.message || 'ファイルロック中です。編集はできません。';
				this.backToList();
			}
		} catch (error) {
			// 一時的な通信失敗では直ちに操作状態を変えず、保存時にサーバーで再検証する
			console.error('ファイルロックの更新に失敗しました。', error);
		}
	},

	releaseFileLock() {
		// 画面を閉じる際も送信できるようsendBeaconを優先する。
		if (this.ignoreFileLock) return;
		if (!this.hasFileLock) return;
		this.hasFileLock = false;
		if (navigator.sendBeacon) {
			navigator.sendBeacon('./php/lock.php?action=release&client=' + encodeURIComponent(this.lockClientId), new Blob([], { type: 'application/octet-stream' }));
		} else {
			fetch('./php/lock.php?action=release&client=' + encodeURIComponent(this.lockClientId), { method: 'POST', keepalive: true });
		}
	}
};

