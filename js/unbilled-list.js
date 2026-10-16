// 制作番号があり、まだ請求済みではない案件をタブ区切りテキストで出力する。
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

const requiredColumns = ['制作番号', '案件名', '取引先名', '請求フラグ', 'データ入力日'];

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

	downloadUnbilledList() {
		const indexes = Object.fromEntries(requiredColumns.map(label => [label, this.projectFieldIndex(label)]));
		const missingColumns = requiredColumns.filter(label => indexes[label] < 0);
		if (missingColumns.length) {
			this.message = '未請求リストを出力できません。Excelに必要な列がありません：' + missingColumns.join('、');
			this.messageType = 'error';
			return;
		}

		const today = new Date();
		today.setHours(0, 0, 0, 0);
		const oneYearAgo = new Date(today);
		oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
		const items = this.rows.slice(1).filter(row => {
			const productionNumber = this.unbilledListCell(row[indexes['制作番号']]);
			const billingFlag = this.unbilledListCell(row[indexes['請求フラグ']]);
			const inputDate = this.unbilledListDate(row[indexes['データ入力日']]);
			return productionNumber !== '' && billingFlag !== '請求済み' &&
				inputDate !== null && inputDate >= oneYearAgo && inputDate <= today;
		}).sort((a, b) => {
			const aNumber = this.unbilledListCell(a[indexes['制作番号']]);
			const bNumber = this.unbilledListCell(b[indexes['制作番号']]);
			return bNumber.localeCompare(aNumber, 'ja', { numeric: true, sensitivity: 'base' });
		});
		const lines = [
			['制作番号', '案件名', '会社名'].join('\t'),
			...items.map(row => [
				row[indexes['制作番号']],
				row[indexes['案件名']],
				row[indexes['取引先名']]
			].map(value => this.unbilledListCell(value)).join('\t'))
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

		this.message = '未請求リストを出力しました（' + items.length + '件）。';
		this.messageType = 'success';
	}
};
