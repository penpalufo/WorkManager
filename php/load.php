<?php
// NAS上のExcelファイルを、そのままブラウザへ返す
$excelPath = '\\\\192.168.100.4\\disk1\\WorkManager\\WorkManager---test10.xlsx';

header('Cache-Control: no-store, no-cache, must-revalidate');

if (!file_exists($excelPath)) {
	sendError(404, 'Excelファイルが見つかりません。');
}

if (!is_readable($excelPath)) {
	sendError(500, 'Excelファイルを読み込めません。');
}

$fileSize = filesize($excelPath);
if ($fileSize === false) {
	sendError(500, 'Excelファイルのサイズを取得できません。');
}

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: inline; filename="WorkManager.xlsx"');
header('Content-Length: ' . $fileSize);

if (readfile($excelPath) === false) {
	// 出力開始後はHTTPレスポンスをJSONへ変更できないためログへ記録する
	error_log('WorkManager.xlsx の読み込み中にエラーが発生しました。');
}

exit;

// エラー内容をJSONで返す
function sendError($statusCode, $message)
{
	http_response_code($statusCode);
	header('Content-Type: application/json; charset=UTF-8');
	echo json_encode(
		array('success' => false, 'message' => $message),
		JSON_UNESCAPED_UNICODE
	);
	exit;
}
