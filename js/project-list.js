// 一覧・検索・ページ切り替え（元のExcel行番号を保持）
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

export const projectListComputed = {
headerRow() {
			// Excelの先頭行を画面と各機能で共通の見出しとして扱う。
			return this.rows.length ? this.rows[0] : [];
		},

	listColumnIndexes() {
			// 一覧には業務上必要な4項目だけを、ここで指定した順に表示する。
			return ['制作番号', '取引先名', '案件名', '請求日']
				.map(label => this.headerRow.findIndex(header => String(header).replace(/\s/g, '') === label))
				.filter(index => index >= 0);
		},

filteredEntries() {
			// 表示は新しい案件から並べつつ、編集に必要な元のExcel行番号を保持する。
			const customerIndex = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '取引先名');
			const projectIndex = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '案件名');
			const productionNumberIndex = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '制作番号');
			const productionNumber = this.searchText(this.productionNumberSearch);
			const customer = this.searchText(this.customerSearch);
			const project = this.searchText(this.projectSearch);
			// 制作番号がある案件だけを対象とし、絞り込み後も元のExcel行番号を保持する。
			return this.rows.slice(1).map((row, index) => ({ row, index: index + 1 })).reverse()
				.filter(entry => productionNumberIndex >= 0 && String(entry.row[productionNumberIndex] == null ? '' : entry.row[productionNumberIndex]).trim() !== '' &&
					(!productionNumber || this.searchText(entry.row[productionNumberIndex]).includes(productionNumber)) &&
					(!customer || this.searchText(entry.row[customerIndex]).includes(customer)) &&
					(!project || this.searchText(entry.row[projectIndex]).includes(project)));
		},

dataRows() {
			return this.filteredEntries.map(entry => entry.row);
		},

totalItems() {
			return this.dataRows.length;
		},

totalPages() {
			return Math.max(1, Math.ceil(this.totalItems / this.pageSize));
		},

pageStartIndex() {
			return (this.currentPage - 1) * this.pageSize;
		},

paginatedRows() {
			return this.dataRows.slice(
				this.pageStartIndex,
				this.pageStartIndex + this.pageSize
			);
		}
};

export const projectListWatch = {
// 検索条件が変わったら、検索結果の1ページ目へ戻す。
productionNumberSearch() { this.currentPage = 1; },

customerSearch() { this.currentPage = 1; },

projectSearch() { this.currentPage = 1; },

totalPages(value) { this.currentPage = Math.min(this.currentPage, value); }
};

export const projectListMethods = {
// ページ移動と、列数が不足する行の表示補完を担当する。
previousPage() {
			if (this.currentPage > 1) {
				this.currentPage--;
			}
		},

nextPage() {
			if (this.currentPage < this.totalPages) {
				this.currentPage++;
			}
		},

displayRow(row) {
			return this.listColumnIndexes.map(index => {
				const value = index < row.length ? row[index] : '';
				const label = String(this.headerRow[index] || '').replace(/\s/g, '');
				if (label !== '請求日' || row === this.headerRow) return value;
				// 一覧の請求日は月日を2桁にそろえ、YYYY/MM/DDで表示する。
				const pickerValue = this.datePickerValue(value, '請求日');
				return pickerValue ? pickerValue.replace(/-/g, '/') : value;
			});
		}
};
