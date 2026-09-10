import { APP_VERSION, APP_UPDATED_AT } from './config.js';
import { masterMethods } from './master.js';
import { projectListComputed, projectListWatch, projectListMethods } from './project-list.js';
import { projectFormComputed, projectFormMethods } from './project-form.js';
import { excelMethods } from './excel.js';
import { utilsMethods } from './utils.js';

console.log('ver ' + APP_VERSION);

// VueとSheetJSはindex.htmlで読み込むCDN版を使用する。
// 状態と起動処理をここに置き、機能別のOptions APIを組み合わせる。
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
    ...utilsMethods
},
mounted() {
		this.loadExcel();
		this.loadMasterExcels();
	}
}).mount('#app');
