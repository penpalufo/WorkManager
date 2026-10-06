// NAS上の排他ロックを取得し、編集中は定期的に有効期限を更新する
export const fileLockMethods = {
	async acquireFileLock() {
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
		if (!this.hasFileLock) return;
		this.hasFileLock = false;
		if (navigator.sendBeacon) {
			navigator.sendBeacon('./php/lock.php?action=release&client=' + encodeURIComponent(this.lockClientId), new Blob([], { type: 'application/octet-stream' }));
		} else {
			fetch('./php/lock.php?action=release&client=' + encodeURIComponent(this.lockClientId), { method: 'POST', keepalive: true });
		}
	}
};

