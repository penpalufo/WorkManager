// 月別の請求分を取引先ごとに集計して表示する。
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

const requiredColumns = ['請求日', '制作番号', '案件名', '取引先名', '担当者氏名', '税込請求金額'];

export const billingReportComputed = {
	billingReportTitle() {
		const match = String(this.billingReportMonth || '').match(/^(\d{4})-(\d{2})$/);
		return match ? match[1] + '年' + Number(match[2]) + '月 請求分' : '請求分';
	},

	billingReportMissingColumns() {
		return requiredColumns.filter(label => this.projectFieldIndex(label) < 0);
	},

	billingReportGroups() {
		if (!this.billingReportMonth || this.billingReportMissingColumns.length) return [];
		const indexes = Object.fromEntries(requiredColumns.map(label => [label, this.projectFieldIndex(label)]));
		const groups = new Map();

		this.rows.slice(1).forEach((row, rowOffset) => {
			const date = this.billingReportDate(row[indexes['請求日']]);
			if (!date || date.monthKey !== this.billingReportMonth) return;
			const amount = this.numericBillingValue(row[indexes['税込請求金額']]);
			const customer = String(row[indexes['取引先名']] || '').trim() || '（取引先名なし）';
			const item = {
				rowNumber: rowOffset + 2,
				dateKey: date.dateKey,
				date: date.display,
				productionNumber: String(row[indexes['制作番号']] || ''),
				projectName: String(row[indexes['案件名']] || ''),
				contactName: String(row[indexes['担当者氏名']] || ''),
				amount
			};
			if (!groups.has(customer)) groups.set(customer, []);
			groups.get(customer).push(item);
		});

		return [...groups.entries()].map(([customer, items]) => {
			items.sort((a, b) => a.dateKey.localeCompare(b.dateKey) ||
				a.productionNumber.localeCompare(b.productionNumber, 'ja', { numeric: true }));
			return {
				customer,
				items,
				total: items.reduce((sum, item) => sum + (item.amount === null ? 0 : item.amount), 0)
			};
		}).sort((a, b) => a.customer.localeCompare(b.customer, 'ja'));
	},

	billingReportTotal() {
		return this.billingReportGroups.reduce((sum, group) => sum + group.total, 0);
	},

	billingReportInvalidItems() {
		if (!this.billingReportMonth || this.billingReportMissingColumns.length) return [];
		const dateIndex = this.projectFieldIndex('請求日');
		const amountIndex = this.projectFieldIndex('税込請求金額');
		return this.rows.slice(1).map((row, index) => ({ row, rowNumber: index + 2 }))
			.filter(entry => {
				const date = this.billingReportDate(entry.row[dateIndex]);
				return date && date.monthKey === this.billingReportMonth &&
					this.numericBillingValue(entry.row[amountIndex]) === null;
			});
	}
};

export const billingReportMethods = {
	defaultBillingReportMonth() {
		const now = new Date();
		return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
	},

	openBillingReport() {
		if (!this.billingReportMonth) this.billingReportMonth = this.defaultBillingReportMonth();
		this.viewMode = 'billing-report';
		this.message = '';
		this.messageType = '';
		window.scrollTo({ top: 0, behavior: 'smooth' });
	},

	billingReportDate(value) {
		const text = this.formatEditValue(value, '請求日').trim();
		const match = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s.*)?$/);
		if (!match) return null;
		const year = Number(match[1]);
		const month = Number(match[2]);
		const day = Number(match[3]);
		const date = new Date(year, month - 1, day);
		if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
		return {
			monthKey: year + '-' + String(month).padStart(2, '0'),
			dateKey: year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0'),
			display: year + '/' + String(month).padStart(2, '0') + '/' + String(day).padStart(2, '0')
		};
	},

	formatBillingReportAmount(value) {
		return value === null ? '金額不正' : '￥' + Number(value).toLocaleString('ja-JP');
	},

	printBillingReport() {
		window.print();
	}
};
