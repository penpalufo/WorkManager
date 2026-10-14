// 文字列・日付の表示とHTTP応答の共通処理
// 各メソッドのthisは、app.jsで組み合わせたVueインスタンスを参照します。

export const utilsMethods = {
searchText(value) {
			// 全角・半角や大文字・小文字の差を吸収した検索文字列を作る。
			return String(value == null ? '' : value).normalize('NFKC').trim().toLowerCase();
		},

formatEditValue(value, label) {
			// 金額と日付だけ項目に適した表示へ整え、それ以外は文字列として返す。
			const text = this.formatCell(value);
			if (this.isAmountField && this.isAmountField(label)) return this.formatAmountField(text);
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

datePickerValue(value, label) {
			const text = this.formatEditValue(value, label);
			const match = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
			return match ? match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0') : '';
		},

fieldLabel(value) {
			// Excel由来の長い関連列名を、画面で読みやすい名称へ置き換える。
			const label = String(value);
			const aliases = {
				'取引先担当者住所参照::担当者ID': '担当者ID',
				'取引先名マスタ::取引先ID': '取引先ID'
			};
			return aliases[label.replace(/\s/g, '')] || label;
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
			// JSON以外のサーバー応答も利用者向けエラーへ変換する。
			const text = await response.text();
			try {
				return JSON.parse(text);
			} catch (error) {
				throw new Error(text || 'サーバーから不正な応答が返されました。');
			}
		}
};
