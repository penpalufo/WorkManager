// 制作番号があり、まだ請求済みではない案件を画面表示し、タブ区切りテキストで出力する。
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

const requiredColumns = ['制作番号', '案件名', '取引先名', '請求フラグ', 'データ入力日'];

export const unbilledListComputed = {
	unbilledListMissingColumns() {
		return requiredColumns.filter(label => this.projectFieldIndex(label) < 0);
	},

	unbilledListItems() {
		if (this.unbilledListMissingColumns.length) return [];
		const indexes = Object.fromEntries(requiredColumns.map(label => [label, this.projectFieldIndex(label)]));
		const today = new Date();
		today.setHours(0, 0, 0, 0);
		const oneYearAgo = new Date(today);
		oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

		return this.rows.slice(1).map((row, index) => ({ row, rowNumber: index + 2 })).filter(entry => {
			const productionNumber = this.unbilledListCell(entry.row[indexes['制作番号']]);
			const billingFlag = this.unbilledListCell(entry.row[indexes['請求フラグ']]);
			const inputDate = this.unbilledListDate(entry.row[indexes['データ入力日']]);
			return productionNumber !== '' && billingFlag !== '請求済み' &&
				inputDate !== null && inputDate >= oneYearAgo && inputDate <= today;
		}).map(entry => ({
			rowNumber: entry.rowNumber,
			productionNumber: this.unbilledListCell(entry.row[indexes['制作番号']]),
			projectName: this.unbilledListCell(entry.row[indexes['案件名']]),
			companyName: this.unbilledListCell(entry.row[indexes['取引先名']])
		})).sort((a, b) => b.productionNumber.localeCompare(
			a.productionNumber,
			'ja',
			{ numeric: true, sensitivity: 'base' }
		));
	}
};

export const unbilledListMethods = {
	// タブや改行を空白へ置き換え、1つの値がTSVの列や行を分断しないようにする。
	unbilledListCell(value) {
		return String(value == null ? '' : value).replace(/[\t\r\n]+/g, ' ').trim();
	},

	// 案件のデータ入力日を、時刻の影響を受けないローカル日付へ変換する。
	unbilledListDate(value) {
		const pickerValue = this.datePickerValue(value, 'データ入力日');
		const match = pickerValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
		return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : null;
	},

	openUnbilledList() {
		this.viewMode = 'unbilled-list';
		this.message = '';
		this.messageType = '';
		window.scrollTo({ top: 0, behavior: 'smooth' });
	},

	downloadUnbilledList() {
		if (this.unbilledListMissingColumns.length) return;
		const lines = [
			['制作番号', '案件名', '会社名'].join('\t'),
			...this.unbilledListItems.map(item => [
				item.productionNumber,
				item.projectName,
				item.companyName
			].join('\t'))
		];
		const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/tab-separated-values;charset=utf-8' });
		const url = URL.createObjectURL(blob);
		const link = document.createElement('a');
		const now = new Date();
		const date = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
		link.href = url;
		link.download = '未請求リスト_' + date + '.txt';
		document.body.appendChild(link);
		link.click();
		link.remove();
		URL.revokeObjectURL(url);

		this.message = '未請求リストを出力しました（' + this.unbilledListItems.length + '件）。';
		this.messageType = 'success';
	}
};
