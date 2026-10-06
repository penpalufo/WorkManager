// 編集中の案件から、A4横1ページの請求書PDFを生成する
export const invoicePdfMethods = {
	invoiceValue(label) {
		const normalized = String(label).replace(/\s/g, '');
		const index = this.headerRow.findIndex((header) =>
			this.fieldLabel(header).replace(/\s/g, '') === normalized);
		return index < 0 ? '' : this.editableRow[index];
	},

	invoiceAmount(value) {
		const number = Number(String(value == null ? '' : value).replace(/[￥¥,\s]/g, ''));
		if (!Number.isFinite(number)) throw new Error('請求金額を数値として読み取れません。');
		return Math.round(number);
	},

	invoiceDate(value) {
		const text = this.formatEditValue(value, '請求日');
		const match = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
		return match ? match[1] + '年' + match[2].padStart(2, '0') + '月' + match[3].padStart(2, '0') + '日' : text;
	},

	async loadInvoiceImage(url) {
		return new Promise((resolve, reject) => {
			const image = new Image();
			image.onload = () => resolve(image);
			image.onerror = () => reject(new Error('社印画像を読み込めませんでした。'));
			image.src = url;
		});
	},

	fitInvoiceText(context, text, maxWidth) {
		let value = String(text || '');
		if (context.measureText(value).width <= maxWidth) return value;
		while (value.length && context.measureText(value + '…').width > maxWidth) value = value.slice(0, -1);
		return value + '…';
	},

	async createInvoicePdf() {
		if (this.viewMode !== 'edit' || this.isGeneratingPdf) return;
		let printWindow = null;
		this.isGeneratingPdf = true;
		this.message = '';
		this.messageType = '';

		try {
			printWindow = window.open('', '_blank', 'width=1200,height=850');
			if (!printWindow) throw new Error('印刷画面を開けませんでした。ポップアップを許可してください。');
			printWindow.document.write('<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>請求書</title></head><body>印刷画面を準備しています...</body></html>');
			printWindow.document.close();
			const customerName = String(this.invoiceValue('取引先名') || '').trim();
			const projectName = String(this.invoiceValue('案件名') || '').trim();
			const productionNumber = String(this.invoiceValue('制作番号') || '').trim();
			const billedAt = this.invoiceDate(this.invoiceValue('請求日'));
			const subtotal = this.invoiceAmount(this.invoiceValue('請求金額'));
			const tax = this.invoiceAmount(this.invoiceValue('税額'));
			const total = this.invoiceAmount(this.invoiceValue('税込請求金額'));
			if (!customerName || !projectName || !billedAt) {
				throw new Error('取引先名、案件名、請求日を入力してください。');
			}

			const yen = (amount) => '￥' + amount.toLocaleString('ja-JP');
			const canvas = document.createElement('canvas');
			canvas.width = 1684;
			canvas.height = 1190;
			const context = canvas.getContext('2d');
			const blue = '#304d83';
			const dark = '#2d2d2d';
			context.fillStyle = '#ffffff';
			context.fillRect(0, 0, canvas.width, canvas.height);
			context.textBaseline = 'alphabetic';

			context.fillStyle = '#686868';
			context.font = 'bold 48px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText('御請求書', 80, 120);

			context.fillStyle = dark;
			context.font = 'bold 29px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText(this.fitInvoiceText(context, customerName + ' 御中', 880), 80, 180);
			context.font = '20px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText(this.fitInvoiceText(context, '制作案件名：  ' + projectName, 900), 80, 225);

			context.textAlign = 'right';
			context.font = '17px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText(billedAt, 1595, 95);
			context.font = '13px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText(productionNumber, 1595, 122);

			context.textAlign = 'left';
			context.font = '42px Georgia, serif';
			context.fillText('Chaordic Design', 1075, 232);
			context.fillStyle = '#503068';
			context.beginPath(); context.arc(1422, 215, 20, 0, Math.PI * 2); context.fill();
			context.fillStyle = '#8ed6e4';
			context.beginPath(); context.arc(1448, 215, 20, 0, Math.PI * 2); context.fill();
			context.fillStyle = dark;
			context.textAlign = 'right';
			context.font = '15px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText('東京都中央区日本橋蛎殻町 1-25-4 日本橋栄ビル 5F', 1452, 267);
			context.fillText('TEL.03-5623-3775　FAX.03-5623-3776', 1452, 291);
			context.fillText('株式会社カオディックデザイン', 1452, 315);
			const stamp = await this.loadInvoiceImage('./company-stamp.jpg');
			context.drawImage(stamp, 1460, 185, 132, 132);

			context.textAlign = 'left';
			context.font = '22px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText('下記の通り、御請求申し上げます。', 80, 315);
			context.fillStyle = blue;
			context.font = 'bold 30px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText('請求金額(税込)', 80, 385);
			context.textAlign = 'right';
			context.fillText(yen(total), 555, 385);
			context.strokeStyle = blue;
			context.lineWidth = 2;
			context.beginPath(); context.moveTo(78, 395); context.lineTo(560, 395); context.stroke();

			const left = 78;
			const right = 1602;
			const tableTop = 420;
			context.beginPath(); context.moveTo(left, tableTop); context.lineTo(right, tableTop); context.stroke();
			context.font = 'bold 19px "Yu Gothic", "Meiryo", sans-serif';
			context.textAlign = 'center';
			context.fillText('細目(商品)名', 505, 455);
			context.fillText('単価', 1220, 455);
			context.fillText('数量', 1345, 455);
			context.fillText('金額', 1530, 455);
			context.beginPath(); context.moveTo(left, 469); context.lineTo(right, 469); context.stroke();
			context.font = '17px "Yu Gothic", "Meiryo", sans-serif';
			context.textAlign = 'left';
			context.fillText(this.fitInvoiceText(context, projectName, 940), 80, 497);
			context.textAlign = 'right';
			context.fillText(yen(total), 1265, 497);
			context.textAlign = 'center';
			context.fillText('一式', 1345, 497);
			context.textAlign = 'right';
			context.fillText(yen(total), 1588, 497);

			const summaryTop = 875;
			context.strokeStyle = blue;
			context.fillStyle = dark;
			context.textAlign = 'left';
			context.font = '18px "Yu Gothic", "Meiryo", sans-serif';
			[['小計', subtotal], ['消費税', tax], ['合計', total]].forEach((entry, index) => {
				const y = summaryTop + index * 48;
				context.beginPath(); context.moveTo(left, y); context.lineTo(right, y); context.stroke();
				context.fillText(entry[0], 80, y + 32);
				context.textAlign = 'right';
				context.fillText(yen(entry[1]), 1588, y + 32);
				context.textAlign = 'left';
			});
			context.beginPath(); context.moveTo(left, summaryTop + 144); context.lineTo(right, summaryTop + 144); context.stroke();

			context.fillStyle = '#111111';
			context.font = '14px "Yu Gothic", "Meiryo", sans-serif';
			context.fillText('お振込み先／三菱UFJ銀行　大伝馬町支店　普通2015615　株式会社カオディックデザイン', 80, 1060);
			context.fillText('登録番号／T8010001074381', 80, 1085);

			const imageUrl = canvas.toDataURL('image/png');
			printWindow.document.open();
			printWindow.document.write('<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>請求書</title><style>@page{size:A4 landscape;margin:0}html,body{width:297mm;height:210mm;margin:0;padding:0;background:#fff}img{display:block;width:297mm;height:210mm;object-fit:contain}@media screen{body{margin:auto;box-shadow:0 0 12px #999}}</style></head><body onload="setTimeout(function(){window.print()},200)"><img src="' + imageUrl + '" alt="請求書"></body></html>');
			printWindow.document.close();
			printWindow.focus();
			this.message = '請求書の印刷画面を開きました';
			this.messageType = 'success';
		} catch (error) {
			if (printWindow && !printWindow.closed) printWindow.close();
			console.error(error);
			this.message = 'PDFの作成に失敗しました: ' + error.message;
			this.messageType = 'error';
		} finally {
			this.isGeneratingPdf = false;
		}
	}
};

