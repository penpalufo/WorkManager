// WorkManagerの起点。機能別モジュールを1つのVueアプリへまとめる。
import { APP_VERSION, APP_UPDATED_AT } from './config.js';
import { masterMethods } from './master.js';
import { projectListComputed, projectListWatch, projectListMethods } from './project-list.js';
import { projectFormComputed, projectFormMethods } from './project-form.js';
import { excelMethods } from './excel.js';
import { utilsMethods } from './utils.js';
import { fileLockMethods } from './file-lock.js';
import { invoicePdfMethods } from './invoice-pdf.js';
import { billingReportComputed, billingReportMethods } from './billing-report.js';

console.log('ver ' + APP_VERSION);

// VueとSheetJSはindex.htmlで読み込むCDN版を使用する。
// 状態と起動処理をここに置き、機能別のOptions APIを組み合わせる。
Vue.createApp({
data() {
		return {
			// 同じブラウザが取得したロックかどうかをサーバー側で識別するID
			lockClientId: window.crypto && window.crypto.randomUUID ? window.crypto.randomUUID() : String(Date.now()) + '-' + Math.random(),
			// URLに?ignoreLock=1がある場合だけ、保守用としてファイルロックを迂回する
			ignoreFileLock: new URLSearchParams(window.location.search).get('ignoreLock') === '1',
			// 将来再表示しやすいよう、「その他」はデータを残したまま表示だけ切り替える
			showOtherSection: false,
			// 画面表示とExcelデータの基本状態
			appVersion: APP_VERSION,
			appUpdatedAt: APP_UPDATED_AT,
			workbook: null,
			sheetName: '',
			rows: [],
			columnCount: 0,
			currentPage: 1,
			pageSize: 100,
			productionNumberSearch: '',
			customerSearch: '',
			projectSearch: '',
			billingReportMonth: '',
			viewMode: 'list',
			selectedRowIndex: -1,
			editableRow: [],
			// 取引先名マスターと取引先担当者マスター
			customerMaster: null,
			customerContactMaster: null,
			isLoading: false,
			isMasterLoading: false,
			isSaving: false,
			isGeneratingPdf: false,
			// 利用者へ表示する処理結果とファイルロック状態
			message: '',
			messageType: '',
			masterLoadMessage: '',
			hasFileLock: false,
			isFileLocked: false,
			lockMessage: '',
			lockRefreshTimer: null
		};
	},
computed: {
	// 各画面で使用する算出値を機能別モジュールから取り込む
    ...projectListComputed,
    ...projectFormComputed,
    ...billingReportComputed
},
watch: {
    ...projectListWatch
},
methods: {
	// 機能別に分けた処理をVueテンプレートから呼び出せるようにする
    ...masterMethods,
    ...projectListMethods,
    ...projectFormMethods,
    ...excelMethods,
    ...utilsMethods,
    ...fileLockMethods,
    ...invoicePdfMethods,
    ...billingReportMethods
},
async mounted() {
		// 編集ロックを確保してから、案件Excelとマスターを読み込む
		await this.acquireFileLock();
		this.lockRefreshTimer = window.setInterval(() => this.refreshFileLock(), 60000);
		window.addEventListener('beforeunload', this.releaseFileLock);
		this.loadExcel();
		this.loadMasterExcels();
	},
beforeUnmount() {
		window.clearInterval(this.lockRefreshTimer);
		window.removeEventListener('beforeunload', this.releaseFileLock);
		this.releaseFileLock();
	}
}).mount('#app');
