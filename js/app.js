import { APP_VERSION, APP_UPDATED_AT } from './config.js';
import { masterMethods } from './master.js';
import { projectListComputed, projectListWatch, projectListMethods } from './project-list.js';
import { projectFormComputed, projectFormMethods } from './project-form.js';
import { excelMethods } from './excel.js';
import { utilsMethods } from './utils.js';
import { fileLockMethods } from './file-lock.js';
import { invoicePdfMethods } from './invoice-pdf.js';

console.log('ver ' + APP_VERSION);

// VueとSheetJSはindex.htmlで読み込むCDN版を使用する。
// 状態と起動処理をここに置き、機能別のOptions APIを組み合わせる。
Vue.createApp({
data() {
		return {
			lockClientId: window.crypto && window.crypto.randomUUID ? window.crypto.randomUUID() : String(Date.now()) + '-' + Math.random(),
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
			isGeneratingPdf: false,
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
    ...projectListComputed,
    ...projectFormComputed
},
watch: {
    ...projectListWatch
},
methods: {
    ...masterMethods,
    ...projectListMethods,
    ...projectFormMethods,
    ...excelMethods,
    ...utilsMethods,
    ...fileLockMethods,
    ...invoicePdfMethods
},
async mounted() {
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
