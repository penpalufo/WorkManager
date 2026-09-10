// 取引先・担当者マスターの取得、重複除去、JSON保存
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

export const masterMethods = {
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
		}
};
