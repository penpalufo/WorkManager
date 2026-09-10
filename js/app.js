'use strict';
// 更新時はこのバージョン・日付とCHANGELOG.mdを揃える
const APP_VERSION = '0.3.0';
const APP_UPDATED_AT = '2026/9/10';
console.log('ver ' + APP_VERSION);

Vue.createApp({
	data() {
		return {
			appVersion: APP_VERSION,
			appUpdatedAt: APP_UPDATED_AT,
			workbook: null,
			sheetName: '',
			rows: [],
			columnCount: 0,
			currentPage: 1,
			pageSize: 100,
			customerSearch: '',
			projectSearch: '',
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
		// JSONの取引先名を選択候補にする
		customerNames() {
			const rows = this.customerMaster ? this.customerMaster.rows : [];
			return [...new Set(rows.map(row => String(row['取引先名'] || '')).filter(name => name.trim()))];
		},
		selectedCustomerName() {
			const index = this.headerRow.findIndex(label => String(label).replace(/\s/g, '') === '取引先名');
			return index < 0 ? '' : String(this.editableRow[index] || '').trim();
		},

		// 選択中の取引先と一致する担当者氏名だけを候補にする
		contactNames() {
			if (!this.customerContactMaster || !this.selectedCustomerName) return [];
			return [...new Set(this.customerContactMaster.rows
				.filter(row => String(row['取引先名'] || '').trim() === this.selectedCustomerName)
				.map(row => String(row['担当者氏名'] || '').trim())
				.filter(Boolean))];
		},

		// 1行目は全ページ共通の見出しとして扱う
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
				const label = this.fieldLabel(header || '項目' + (index + 1));
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

	watch: {
		customerSearch() { this.currentPage = 1; },
		projectSearch() { this.currentPage = 1; },
		totalPages(value) { this.currentPage = Math.min(this.currentPage, value); }
	},

	methods: {
		// セルの値・書式を保持しながら削除行より下を1行上へ移す
		worksheetWithoutRow(source, rowIndex) {
			const sheet = { ...source };
			for (const address of Object.keys(source)) {
				if (!/^[A-Z]+[1-9]\d*$/.test(address)) continue;
				const cell = XLSX.utils.decode_cell(address);
				if (cell.r >= rowIndex) delete sheet[address];
			}
			for (const address of Object.keys(source)) {
				if (!/^[A-Z]+[1-9]\d*$/.test(address)) continue;
				const cell = XLSX.utils.decode_cell(address);
				if (cell.r > rowIndex) sheet[XLSX.utils.encode_cell({ r: cell.r - 1, c: cell.c })] = source[address];
			}
			const range = XLSX.utils.decode_range(source['!ref']);
			range.e.r = Math.max(range.s.r, range.e.r - 1);
			sheet['!ref'] = XLSX.utils.encode_range(range);
			if (source['!rows']) {
				sheet['!rows'] = source['!rows'].slice();
				sheet['!rows'].splice(rowIndex, 1);
			}
			if (source['!merges']) {
				sheet['!merges'] = source['!merges'].filter(m => !(m.s.r === rowIndex && m.e.r === rowIndex))
					.map(m => ({ s: { c: m.s.c, r: m.s.r > rowIndex ? m.s.r - 1 : m.s.r },
						e: { c: m.e.c, r: m.e.r >= rowIndex ? m.e.r - 1 : m.e.r } }));
			}
			if (source['!autofilter'] && source['!autofilter'].ref) {
				const filter = XLSX.utils.decode_range(source['!autofilter'].ref);
				if (filter.e.r >= rowIndex) filter.e.r = Math.max(filter.s.r, filter.e.r - 1);
				sheet['!autofilter'] = { ...source['!autofilter'], ref: XLSX.utils.encode_range(filter) };
			}
			return sheet;
		},

		async deleteProject() {
			if (this.isSaving || this.viewMode !== 'edit' || this.selectedRowIndex < 1 || !this.workbook) return;
			const rowIndex = this.selectedRowIndex;
			const nameIndex = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '案件名');
			const name = this.rows[rowIndex][nameIndex] || '選択中の案件';
			if (!window.confirm('「' + name + '」を削除しますか？\nExcelからこの案件を削除します。この操作は取り消せません。')) return;
			this.isSaving = true;
			this.message = '';
			try {
				const sheet = this.worksheetWithoutRow(this.workbook.Sheets[this.sheetName], rowIndex);
				const candidate = { ...this.workbook, Sheets: { ...this.workbook.Sheets, [this.sheetName]: sheet } };
				const response = await fetch('./php/save.php', {
					method: 'POST', headers: { 'Content-Type': 'application/octet-stream' },
					body: XLSX.write(candidate, { bookType: 'xlsx', type: 'array' })
				});
				const result = await this.getJsonResponse(response);
				if (!response.ok || !result.success) throw new Error(result.message || '削除を保存できませんでした。');
				this.workbook = candidate;
				this.rows.splice(rowIndex, 1);
				this.selectedRowIndex = -1;
				this.editableRow = [];
				this.viewMode = 'list';
				this.message = '案件を削除しました';
				this.messageType = 'success';
			} catch (error) {
				this.message = '削除に失敗しました: ' + error.message;
				this.messageType = 'error';
			} finally {
				this.isSaving = false;
			}
		},

		// 表示順や検索条件ではなく、Excel末尾の制作番号を基準にする
		nextProductionNumber(index) {
			for (let r = this.rows.length - 1; r >= 1; r--) {
				const value = String(this.rows[r][index] == null ? '' : this.rows[r][index]).trim().replace(/,/g, '');
				if (!value) continue;
				if (!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value) + 1)) {
					throw new Error('最後の制作番号を数値として採番できません。');
				}
				return String(Number(value) + 1).padStart(value.length, '0');
			}
			return '1';
		},

		openNewProject() {
			if (!this.workbook || this.isSaving) return;
			try {
				const index = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '制作番号');
				if (index < 0) throw new Error('Excelに制作番号列がありません。');
				this.editableRow = Array(this.columnCount).fill('');
				this.editableRow[index] = this.nextProductionNumber(index);
				this.selectedRowIndex = -1;
				this.viewMode = 'new';
				this.message = '';
				this.messageType = '';
				window.scrollTo({ top: 0, behavior: 'smooth' });
			} catch (error) {
				this.message = error.message;
				this.messageType = 'error';
			}
		},

		async addProject() {
			if (this.isSaving || !this.workbook || this.viewMode !== 'new') return;
			this.isSaving = true;
			this.message = '';
			try {
				const index = this.headerRow.findIndex(h => String(h).replace(/\s/g, '') === '制作番号');
				if (index < 0) throw new Error('Excelに制作番号列がありません。');
				const row = this.editableRow.slice();
				row[index] = this.nextProductionNumber(index);
				// 成功するまでは元のWorkbookに追加行を残さない
				const worksheet = { ...this.workbook.Sheets[this.sheetName] };
				const candidate = { ...this.workbook, Sheets: { ...this.workbook.Sheets, [this.sheetName]: worksheet } };
				XLSX.utils.sheet_add_aoa(worksheet, [row], { origin: -1 });
				const response = await fetch('./php/save.php', {
					method: 'POST', headers: { 'Content-Type': 'application/octet-stream' },
					body: XLSX.write(candidate, { bookType: 'xlsx', type: 'array' })
				});
				const result = await this.getJsonResponse(response);
				if (!response.ok || !result.success) throw new Error(result.message || '案件を追加できませんでした。');
				this.workbook = candidate;
				this.rows.push(row);
				this.customerSearch = '';
				this.projectSearch = '';
				this.currentPage = 1;
				this.viewMode = 'list';
				this.editableRow = [];
				this.message = '案件を追加しました';
				this.messageType = 'success';
			} catch (error) {
				this.message = '追加に失敗しました: ' + error.message;
				this.messageType = 'error';
			} finally {
				this.isSaving = false;
			}
		},
		searchText(value) {
			return String(value == null ? '' : value).normalize('NFKC').trim().toLowerCase();
		},
		// 指定の日付項目だけを年/月/日に整える（元データは変更しない）
		formatEditValue(value, label) {
			const text = this.formatCell(value);
			if (!['請求日', '請求予定日', 'データ入力日', '集計日'].includes(String(label).replace(/\s/g, ''))) return text;
			const match = text.trim().match(/^(\d{1,4})\/(\d{1,2})\/(\d{1,4})$/);
			if (!match) return text;
			let year, month, day;
			if (match[1].length === 4) {
				year = Number(match[1]); month = Number(match[2]); day = Number(match[3]);
			} else {
				month = Number(match[1]); day = Number(match[2]); year = Number(match[3]);
				if (match[3].length <= 2) year += year < 30 ? 2000 : 1900;
			}
			const date = new Date(year, month - 1, day);
			if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return text;
			return year + '/' + month + '/' + day;
		},

		// カレンダー入力にはゼロ埋めしたYYYY-MM-DDを渡す
		datePickerValue(value, label) {
			const text = this.formatEditValue(value, label);
			const match = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
			return match ? match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0') : '';
		},

		setPickedDate(index, value) {
			if (!value) {
				this.editableRow[index] = '';
				return;
			}
			const parts = value.split('-');
			this.editableRow[index] = parts[0] + '/' + Number(parts[1]) + '/' + Number(parts[2]);
		},

		// Excelの元の列名は保持し、表示と列判定に共通の名前を使う
		fieldLabel(value) {
			const label = String(value);
			const aliases = {
				'取引先担当者住所参照::担当者ID': '担当者ID',
				'取引先名マスタ::取引先ID': '取引先ID'
			};
			return aliases[label.replace(/\s/g, '')] || label;
		},

		// 取引先を変更したら以前の担当者を解除する
		onCustomerChanged() {
			const customer = this.customerMaster && this.customerMaster.rows.find(row =>
				String(row['取引先名'] || '').trim() === this.selectedCustomerName);
			this.headerRow.forEach((label, index) => {
				if (this.fieldLabel(label).replace(/\s/g, '') === '取引先ID') {
					this.editableRow[index] = customer ? customer['取引先ID'] || '' : '';
				}
				if (['担当者氏名', '担当者ID'].includes(this.fieldLabel(label).replace(/\s/g, ''))) {
					this.editableRow[index] = '';
				}
			});
		},

		// 選択した取引先・担当者氏名に対応する担当者IDを反映する
		onContactChanged() {
			const nameIndex = this.headerRow.findIndex(label =>
				String(label).replace(/\s/g, '') === '担当者氏名');
			const name = nameIndex < 0 ? '' : String(this.editableRow[nameIndex] || '').trim();
			const contact = name && this.customerContactMaster && this.customerContactMaster.rows.find(row =>
				String(row['取引先名'] || '').trim() === this.selectedCustomerName &&
				String(row['担当者氏名'] || '').trim() === name);
			this.headerRow.forEach((label, index) => {
				if (this.fieldLabel(label).replace(/\s/g, '') === '担当者ID') {
					this.editableRow[index] = contact ? contact['担当者ID'] || '' : '';
				}
			});
		},
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
			const seenContacts = new Set();
			for (let index = 1; index < excel.rows.length; index++) {
				const row = excel.rows[index];
				const id = String(row[idIndex] == null ? '' : row[idIndex]).trim();
				if (id && !seenContacts.has(id)) {
					seenContacts.add(id);
					rows.push(row);
				}
			}
			const master = {
				sheetName: 'customerContact',
				rows: XLSX.utils.sheet_to_json(XLSX.utils.aoa_to_sheet(rows), { defval: '', raw: false })
			};
			console.log('担当者マスター: JSON保存開始', master.rows.length, '件');
			const uploadId = Array.from(crypto.getRandomValues(new Uint8Array(16)), v => v.toString(16).padStart(2, '0')).join('');
			let offset = 0;
			let start = 0;
			do {
				const batch = [];
				let bytes = 0;
				while (start < master.rows.length && batch.length < 500) {
					const row = master.rows[start];
					const size = new TextEncoder().encode(JSON.stringify(row)).length + 1;
					if (size > 1000000) throw new Error('担当者1件のデータが1MBを超えています。');
					if (bytes + size > 1000000) break;
					batch.push(row); bytes += size; start++;
				}
			const saved = await fetch('./php/save-customer-contact-master.php', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json' },
				body: JSON.stringify({ id: uploadId, offset: offset, final: start === master.rows.length, rows: batch })
			});
			const result = await this.getJsonResponse(saved);
			if (!saved.ok || !result.success) {
				throw new Error(result.message || '担当者マスターのJSON保存に失敗しました。');
			}
				offset = result.offset;
				console.log('担当者マスター: 保存済み', start, '/', master.rows.length);
			} while (start < master.rows.length);
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
			const entry = this.filteredEntries[this.pageStartIndex + rowIndex];
			if (!entry) return;
			this.selectedRowIndex = entry.index;
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
			if (this.isSaving) return;
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
