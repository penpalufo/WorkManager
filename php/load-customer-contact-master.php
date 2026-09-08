<?php
// NAS上の取引先担当者マスターを、そのままブラウザへ返す
$excelPath = '\\\\192.168.100.4\\disk1\\WorkManager\\取引先担当者マスタ.xlsx';

header('Cache-Control: no-store, no-cache, must-revalidate');

if (!file_exists($excelPath)) {
	sendError(404, '取引先担当者マスターが見つかりません。');
}

if (!is_readable($excelPath)) {
	sendError(500, '取引先担当者マスターを読み込めません。');
}

$fileSize = filesize($excelPath);
if ($fileSize === false) {
	sendError(500, '取引先担当者マスターのサイズを取得できません。');
}

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: inline; filename="customer-contact-master.xlsx"');
header('Content-Length: ' . $fileSize);

if (readfile($excelPath) === false) {
	error_log('取引先担当者マスターの読み込み中にエラーが発生しました。');
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
