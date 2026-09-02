'use strict';

Vue.createApp({
	data() {
		return {
			workbook: null,
			sheetName: '',
			rows: [],
			columnCount: 0,
			isLoading: false,
			isSaving: false,
			message: '',
			messageType: ''
		};
	},

	methods: {
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
			} catch (error) {
				console.error(error);
				this.workbook = null;
				this.rows = [];
				this.columnCount = 0;
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
	}
}).mount('#app');
