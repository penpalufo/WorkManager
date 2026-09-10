// SheetJSによるExcelの取得・解析とワークシート操作
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

export const excelMethods = {
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
		}
};
