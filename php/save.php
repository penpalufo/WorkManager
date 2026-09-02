<?php
// SheetJSが生成したExcelバイナリをNASへ保存する
$excelPath = '\\\\192.168.100.4\\disk1\\WorkManager\\WorkManager---test10.xlsx';

header('Content-Type: application/json; charset=UTF-8');
header('Cache-Control: no-store, no-cache, must-revalidate');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
	sendJson(405, false, 'POSTメソッドで送信してください。');
}

if (!file_exists($excelPath)) {
	sendJson(404, false, '保存先のExcelファイルが見つかりません。');
}

if (!is_writable($excelPath)) {
	sendJson(500, false, 'Excelファイルへ書き込めません。');
}

$data = file_get_contents('php://input');
if ($data === false || strlen($data) === 0) {
	sendJson(400, false, '送信されたExcelデータが空です。');
}

// ExcelファイルはZIP形式なので、先頭が「PK」であることを最低限確認する
if (substr($data, 0, 2) !== 'PK') {
	sendJson(400, false, '送信されたデータは有効なxlsx形式ではありません。');
}

// 元ファイルと同じNASフォルダーへ一時ファイルを作る
$tempPath = $excelPath . '.tmp';
$writtenBytes = @file_put_contents($tempPath, $data, LOCK_EX);

if ($writtenBytes === false || $writtenBytes !== strlen($data)) {
	if (file_exists($tempPath)) {
		@unlink($tempPath);
	}
	sendJson(500, false, '一時ファイルへの書き込みに失敗しました。');
}

clearstatcache(true, $tempPath);
if (@filesize($tempPath) !== strlen($data)) {
	@unlink($tempPath);
	sendJson(500, false, '一時ファイルの書き込み結果を確認できませんでした。');
}

// Windows/NASでは既存ファイルへのrenameが失敗する場合があるため、
// 検証済みの一時ファイルを元ファイルへコピーして反映する
if (!@copy($tempPath, $excelPath)) {
	@unlink($tempPath);
	sendJson(500, false, 'Excelファイルの上書き保存に失敗しました。');
}

@unlink($tempPath);

sendJson(200, true, 'Excelファイルを保存しました。');

// JSON形式で処理結果を返す
function sendJson($statusCode, $success, $message)
{
	http_response_code($statusCode);
	echo json_encode(
		array('success' => $success, 'message' => $message),
		JSON_UNESCAPED_UNICODE
	);
	exit;
}
