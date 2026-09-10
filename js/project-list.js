// 一覧・検索・ページ切り替え（元のExcel行番号を保持）
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

export const projectListComputed = {
headerRow() {
			return this.rows.length ? this.rows[0] : [];
		},

filteredEntries() {
			const customerIndex = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '取引先名');
			const projectIndex = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '案件名');
			const customer = this.searchText(this.customerSearch);
			const project = this.searchText(this.projectSearch);
			// 絞り込み後も元のExcel行番号を保持する
			return this.rows.slice(1).map((row, index) => ({ row, index: index + 1 })).reverse()
				.filter(entry => (!customer || this.searchText(entry.row[customerIndex]).includes(customer)) &&
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
customerSearch() { this.currentPage = 1; },

projectSearch() { this.currentPage = 1; },

totalPages(value) { this.currentPage = Math.min(this.currentPage, value); }
};

export const projectListMethods = {
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
			return Array.from({ length: this.columnCount }, (_, index) => {
				return index < row.length ? row[index] : '';
			});
		}
};
