// 案件の新規追加・編集・削除、取引先と担当者の連動
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

export const projectFormComputed = {
customerNames() {
			const rows = this.customerMaster ? this.customerMaster.rows : [];
			return [...new Set(rows.map(row => String(row['取引先名'] || '')).filter(name => name.trim()))];
		},

selectedCustomerName() {
			const index = this.headerRow.findIndex(label => String(label).replace(/\s/g, '') === '取引先名');
			return index < 0 ? '' : String(this.editableRow[index] || '').trim();
		},

contactNames() {
			if (!this.customerContactMaster || !this.selectedCustomerName) return [];
			return [...new Set(this.customerContactMaster.rows
				.filter(row => String(row['取引先名'] || '').trim() === this.selectedCustomerName)
				.map(row => String(row['担当者氏名'] || '').trim())
				.filter(Boolean))];
		},

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
};

export const projectFormMethods = {
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

setPickedDate(index, value) {
			if (!value) {
				this.editableRow[index] = '';
				return;
			}
			const parts = value.split('-');
			this.editableRow[index] = parts[0] + '/' + Number(parts[1]) + '/' + Number(parts[2]);
		},

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
		}
};
