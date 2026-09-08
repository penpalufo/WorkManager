'use strict';

Vue.createApp({
	data() {
		return {
			workbook: null,
			sheetName: '',
			rows: [],
			columnCount: 0,
			currentPage: 1,
			pageSize: 100,
			viewMode: 'list',
			selectedRowIndex: -1,
			editableRow: [],
			// 取引先名マスターと取引先担当者マスターを格納するオブジェクト
			customerMaster: null,
			customerContactMaster: null,
			isLoading: false,
			isMasterLoading: false,
			isSaving: false,
			message: '',
			messageType: '',
			masterLoadMessage: ''
		};
	},

	computed: {
		// 1行目は全ページ共通の見出しとして扱う
		headerRow() {
			return this.rows.length ? this.rows[0] : [];
		},

		dataRows() {
			// Excelの最終データ行を一覧の先頭に表示する
			return this.rows.length > 1 ? this.rows.slice(1).reverse() : [];
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
		},

		// Excelの見出し名を使い、画像に近いまとまりで編集項目を表示する
		editFieldGroups() {
			const groups = {
				basic: { title: '案件情報', fields: [] },
				customer: { title: '取引先・担当者情報', fields: [] },
				billing: { title: '請求情報', fields: [] },
				other: { title: 'その他', fields: [] }
			};

			this.headerRow.forEach((header, index) => {
				const label = String(header || '項目' + (index + 1));
				const field = { index: index, label: label };

				if (/取引先|顧客|担当者|住所/.test(label)) {
					groups.customer.fields.push(field);
				} else if (/請求|見積|金額|税|入金/.test(label)) {
					groups.billing.fields.push(field);
				} else if (/案件|制作|管理番号|件名/.test(label)) {
					groups.basic.fields.push(field);
				} else {
					groups.other.fields.push(field);
				}
			});

			return Object.values(groups).filter((group) => group.fields.length);
		}
	},

	methods: {
		// 取引先IDのない行を除外し、同じIDは先に登場したレコードを採用する
		uniqueCustomers(master) {
			const headers = master.rows[0] || [];
			const idIndex = headers.findIndex((value) =>
				String(value).normalize('NFKC').replace(/\s/g, '').toLowerCase() === '取引先id');
			if (idIndex < 0) {
				throw new Error('取引先名マスターに「取引先ID」列がありません。');
			}
			const seen = new Set();
			const rows = [headers];
			for (const row of master.rows.slice(1)) {
				const id = String(row[idIndex] == null ? '' : row[idIndex]).trim();
				if (!id || seen.has(id)) continue;
				seen.add(id);
				rows.push(row);
			}
			return {
				sheetName: 'customer',
				rows: XLSX.utils.sheet_to_json(XLSX.utils.aoa_to_sheet(rows), {
					defval: '', raw: false
				})
			};
		},

		async loadCustomerMaster() {
			const response = await fetch('./data/customerMaster.json', { cache: 'no-store' });
			if (response.ok) {
				const master = await response.json();
				if (!master || !Array.isArray(master.rows)) {
					throw new Error('customerMaster.jsonの形式が不正です。');
				}
				// 旧形式のJSONもExcelを取得せずに新形式へ移行する
				if (master.rows.length && Array.isArray(master.rows[0])) {
					const converted = this.uniqueCustomers(master);
					await this.saveCustomerMaster(converted);
					return converted;
				}
				if (master.sheetName !== 'customer' || !master.rows.every((row) =>
					row && typeof row === 'object' && !Array.isArray(row))) {
					throw new Error('customerMaster.jsonの形式が不正です。');
				}
				return master;
			}
			// ファイル未作成の場合だけExcelを読み込む
			if (response.status !== 404) {
				throw new Error('customerMaster.jsonを読み込めません（HTTP ' + response.status + '）。');
			}
			const master = this.uniqueCustomers(await this.loadExcelObject(
				'./php/load-customer-master.php', '取引先名マスター'));
			await this.saveCustomerMaster(master);
			return master;
		},

		async saveCustomerMaster(master) {
			const saved = await fetch('./php/save-customer-master.php', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json' },
				body: JSON.stringify(master)
			});
			const result = await this.getJsonResponse(saved);
			if (!saved.ok || !result.success) {
				throw new Error(result.message || '取引先名マスターのJSON保存に失敗しました。');
			}
		},

		// 担当者マスターもJSONが未作成の場合だけExcelから変換する
		async loadCustomerContactMaster() {
			const response = await fetch('./data/customerContactMaster.json', { cache: 'no-store' });
			if (response.ok) {
				const master = await response.json();
				if (!master || master.sheetName !== 'customerContact' || !Array.isArray(master.rows)
					|| !master.rows.every((row) => row && typeof row === 'object' && !Array.isArray(row))) {
					throw new Error('customerContactMaster.jsonの形式が不正です。');
				}
				return master;
			}
			if (response.status !== 404) {
				throw new Error('customerContactMaster.jsonを読み込めません（HTTP ' + response.status + '）。');
			}
			const excel = await this.loadExcelObject(
				'./php/load-customer-contact-master.php', '取引先担当者マスター', false);
			const headers = excel.rows[0] || [];
			const idIndex = headers.findIndex((value) =>
				String(value).normalize('NFKC').replace(/\s/g, '').toLowerCase() === '担当者id');
			if (idIndex < 0) {
				throw new Error('取引先担当者マスターに「担当者ID」列がありません。');
			}
			// ID未設定・空欄・空白だけの行を除外する
			const rows = [headers];
			for (let index = 1; index < excel.rows.length; index++) {
				const row = excel.rows[index];
				if (String(row[idIndex] == null ? '' : row[idIndex]).trim() !== '') {
					rows.push(row);
				}
			}
			const master = {
				sheetName: 'customerContact',
				rows: XLSX.utils.sheet_to_json(XLSX.utils.aoa_to_sheet(rows), { defval: '', raw: false })
			};
			console.log('担当者マスター: JSON保存開始', master.rows.length, '件');
			const saved = await fetch('./php/save-customer-contact-master.php', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json' },
				body: JSON.stringify(master)
			});
			const result = await this.getJsonResponse(saved);
			if (!saved.ok || !result.success) {
				throw new Error(result.message || '担当者マスターのJSON保存に失敗しました。');
			}
			return master;
		},

		// 指定したPHPからExcelを取得し、扱いやすい1つのオブジェクトにまとめる
		async loadExcelObject(url, label, includeRecords = true) {
			console.log(label + ': Excel取得開始');
			const response = await fetch(url, { cache: 'no-store' });

			if (!response.ok) {
				throw new Error(label + ': ' + await this.getErrorMessage(response));
			}

			const excelBuffer = await response.arrayBuffer();
			if (!excelBuffer.byteLength) {
				throw new Error(label + ': 取得したExcelファイルが空です。');
			}

			console.log(label + ': Excel解析開始', excelBuffer.byteLength, 'bytes');
			const workbook = XLSX.read(excelBuffer, {
				type: 'array',
				cellDates: true
			});

			if (!workbook.SheetNames.length) {
				throw new Error(label + ': Excelファイルにワークシートがありません。');
			}

			const sheetName = workbook.SheetNames[0];
			const worksheet = workbook.Sheets[sheetName];

			console.log(label + ': Excel解析完了・行変換開始');
			return {
				workbook: workbook,
				sheetName: sheetName,
				rows: XLSX.utils.sheet_to_json(worksheet, {
					header: 1,
					defval: '',
					raw: false,
					dateNF: 'yyyy/mm/dd hh:mm:ss'
				}),
				records: includeRecords ? XLSX.utils.sheet_to_json(worksheet, {
					defval: '',
					raw: false,
					dateNF: 'yyyy/mm/dd hh:mm:ss'
				}) : []
			};
		},

		// 2つのマスターExcelを読み込み、それぞれ別のオブジェクトへ格納する
		async loadMasterExcels() {
			this.isMasterLoading = true;
			this.masterLoadMessage = '';

			try {
				const results = await Promise.allSettled([
					this.loadCustomerMaster(),
					this.loadCustomerContactMaster()
				]);

				this.customerMaster = results[0].status === 'fulfilled' ? results[0].value : null;
				this.customerContactMaster = results[1].status === 'fulfilled' ? results[1].value : null;

				console.log('取引先名マスター:', this.customerMaster);
				console.log('取引先担当者マスター:', this.customerContactMaster);

				const errors = results
					.filter((result) => result.status === 'rejected')
					.map((result) => result.reason.message);

				if (errors.length) {
					this.masterLoadMessage = 'マスターファイルの読み込みに失敗しました: ' + errors.join(' / ');
				}
			} catch (error) {
				console.error(error);
				this.masterLoadMessage = 'マスターファイルの読み込みに失敗しました: ' + error.message;
			} finally {
				this.isMasterLoading = false;
			}
		},

		// NAS上のExcelをPHP経由で取得し、最初のシートを読み込む
		async loadExcel() {
			this.isLoading = true;
			this.message = '';
			this.messageType = '';

			try {
				const response = await fetch('./php/load.php', {
					cache: 'no-store'
				});

				if (!response.ok) {
					throw new Error(await this.getErrorMessage(response));
				}

				const excelBuffer = await response.arrayBuffer();
				if (!excelBuffer.byteLength) {
					throw new Error('取得したExcelファイルが空です。');
				}

				// cellDatesを有効にし、日付セルをJavaScriptのDateとして扱う
				const workbook = XLSX.read(excelBuffer, {
					type: 'array',
					cellDates: true
				});

				if (!workbook.SheetNames.length) {
					throw new Error('Excelファイルにワークシートがありません。');
				}

				const sheetName = workbook.SheetNames[0];
				const worksheet = workbook.Sheets[sheetName];
				const rows = XLSX.utils.sheet_to_json(worksheet, {
					header: 1,
					defval: '',
					raw: false,
					dateNF: 'yyyy/mm/dd hh:mm:ss'
				});

				this.workbook = workbook;
				this.sheetName = sheetName;
				this.rows = rows;
				this.columnCount = Math.max(1, ...rows.map((row) => row.length));

				this.currentPage = 1;
			} catch (error) {
				console.error(error);
				this.workbook = null;
				this.rows = [];
				this.columnCount = 0;
				this.currentPage = 1;
				this.message = '読み込みに失敗しました: ' + error.message;
				this.messageType = 'error';
			} finally {
				this.isLoading = false;
			}
		},

		// 現在の列数に合わせたテストデータを作る
		createTestRow() {
			const values = [
				'TEST',
				'テスト案件',
				new Date().toLocaleString('ja-JP'),
				10000,
				'未請求'
			];

			return Array.from({ length: this.columnCount }, (_, index) => {
				return index < values.length ? values[index] : '';
			});
		},

		// テスト行を追加し、再生成したExcelをPHPへ送信する
		async saveTestData() {
			if (!this.workbook || this.isSaving) {
				return;
			}

			this.isSaving = true;
			this.message = '';
			this.messageType = '';

			try {
				const worksheet = this.workbook.Sheets[this.sheetName];
				const newRow = this.createTestRow();

				// 既存シートの末尾へ1行だけ追加する
				XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: -1 });

				const excelData = XLSX.write(this.workbook, {
					bookType: 'xlsx',
					type: 'array'
				});

				const response = await fetch('./php/save.php', {
					method: 'POST',
					headers: {
						'Content-Type': 'application/octet-stream'
					},
					body: excelData
				});

				const result = await this.getJsonResponse(response);
				if (!response.ok || !result.success) {
					throw new Error(result.message || 'Excelファイルを保存できませんでした。');
				}

				this.rows.push(newRow);
				// 逆順表示では追加した最新行が1ページ目に表示される
				this.currentPage = 1;
				this.message = '保存しました';
				this.messageType = 'success';
			} catch (error) {
				console.error(error);
				this.message = '保存に失敗しました: ' + error.message;
				this.messageType = 'error';

				// 失敗時はメモリ上の変更を残さないよう、NASから読み直す
				await this.loadExcel();
				if (!this.message) {
					this.message = '保存に失敗しました。Excelを読み直しました。';
					this.messageType = 'error';
				}
			} finally {
				this.isSaving = false;
			}
		},

		previousPage() {
			if (this.currentPage > 1) {
				this.currentPage--;
			}
		},

		// 一覧で選択した案件を編集画面へ展開する
		openProject(rowIndex) {
			// 逆順の表示位置をExcel上の元の行番号へ戻す
			this.selectedRowIndex = this.totalItems - (this.pageStartIndex + rowIndex);
			this.editableRow = Array.from({ length: this.columnCount }, (_, index) => {
				const value = this.rows[this.selectedRowIndex][index];
				return value === null || typeof value === 'undefined' ? '' : value;
			});
			this.message = '';
			this.messageType = '';
			this.viewMode = 'edit';
			window.scrollTo({ top: 0, behavior: 'smooth' });
		},

		backToList() {
			this.viewMode = 'list';
			this.selectedRowIndex = -1;
			this.editableRow = [];
		},

		// 編集した1行をワークシートへ反映し、NAS上のExcelへ保存する
		async saveProjectChanges() {
			if (!this.workbook || this.selectedRowIndex < 1 || this.isSaving) {
				return;
			}

			this.isSaving = true;
			this.message = '';
			this.messageType = '';
			const originalRow = this.rows[this.selectedRowIndex].slice();

			try {
				const worksheet = this.workbook.Sheets[this.sheetName];
				const updatedRow = this.editableRow.slice();
				XLSX.utils.sheet_add_aoa(worksheet, [updatedRow], {
					origin: { r: this.selectedRowIndex, c: 0 }
				});

				const excelData = XLSX.write(this.workbook, {
					bookType: 'xlsx',
					type: 'array'
				});

				const response = await fetch('./php/save.php', {
					method: 'POST',
					headers: { 'Content-Type': 'application/octet-stream' },
					body: excelData
				});
				const result = await this.getJsonResponse(response);

				if (!response.ok || !result.success) {
					throw new Error(result.message || 'Excelファイルを保存できませんでした。');
				}

				this.rows.splice(this.selectedRowIndex, 1, updatedRow);
				this.message = '変更を保存しました';
				this.messageType = 'success';
			} catch (error) {
				console.error(error);
				// 保存できなかった場合はWorkbookも編集前の値へ戻す
				XLSX.utils.sheet_add_aoa(
					this.workbook.Sheets[this.sheetName],
					[originalRow],
					{ origin: { r: this.selectedRowIndex, c: 0 } }
				);
				this.message = '変更の保存に失敗しました: ' + error.message;
				this.messageType = 'error';
			} finally {
				this.isSaving = false;
			}
		},

		nextPage() {
			if (this.currentPage < this.totalPages) {
				this.currentPage++;
			}
		},

		// 表の行を最大列数まで空セルで補う
		displayRow(row) {
			return Array.from({ length: this.columnCount }, (_, index) => {
				return index < row.length ? row[index] : '';
			});
		},

		formatCell(cell) {
			if (cell === null || typeof cell === 'undefined') {
				return '';
			}
			return String(cell);
		},

		async getErrorMessage(response) {
			const text = await response.text();
			try {
				const data = JSON.parse(text);
				return data.message || 'Excelデータの取得に失敗しました。';
			} catch (error) {
				return text || 'Excelデータの取得に失敗しました。';
			}
		},

		async getJsonResponse(response) {
			const text = await response.text();
			try {
				return JSON.parse(text);
			} catch (error) {
				throw new Error(text || 'サーバーから不正な応答が返されました。');
			}
		}
	},

	mounted() {
		this.loadExcel();
		this.loadMasterExcels();
	}
}).mount('#app');
