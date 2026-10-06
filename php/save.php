<?php
require_once __DIR__ . '/lock-common.php';

// SheetJSが生成したExcelバイナリをNASへ保存する
$excelPath = WORK_MANAGER_EXCEL_PATH;
$backupDirectory = WORK_MANAGER_NAS_DIRECTORY . '\\BACKUP';
$backupLimit = 5;
date_default_timezone_set('Asia/Tokyo');

header('Content-Type: application/json; charset=UTF-8');
header('Cache-Control: no-store, no-cache, must-revalidate');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
	sendJson(405, false, 'POSTメソッドで送信してください。');
}

startWorkManagerSession();
if (!lockClientId() || !sessionOwnsLock(lockClientId())) {
	sendJson(409, false, 'ファイルロック中です。編集はできません。');
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

if (!createBackup($excelPath, $backupDirectory, $backupLimit)) {
	sendJson(500, false, 'バックアップを作成できなかったため、保存を中止しました。');
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

if (!refreshOwnedLock(lockClientId())) {
	sendJson(200, true, 'Excelファイルを保存しましたが、ファイルロックを更新できませんでした。');
}

sendJson(200, true, 'Excelファイルを保存しました。');

function createBackup($sourcePath, $directory, $limit)
{
	if (!is_dir($directory) && !@mkdir($directory, 0777, true)) {
		return false;
	}

	// Windows/NASで使用できないコロンを避け、時刻はハイフン区切りにする
	$backupPath = $directory . '\\' . date('Y-m-d_H-i-s') . '.xls';
	$suffix = 1;
	while (file_exists($backupPath)) {
		$backupPath = $directory . '\\' . date('Y-m-d_H-i-s') . '_' . $suffix . '.xls';
		$suffix++;
	}

	if (!@copy($sourcePath, $backupPath)) {
		return false;
	}

	$files = glob($directory . '\\*.xls');
	if ($files === false) {
		return true;
	}

	usort($files, function ($left, $right) {
		return filemtime($right) <=> filemtime($left);
	});
	foreach (array_slice($files, max(1, $limit)) as $oldBackup) {
		if (!@unlink($oldBackup)) {
			return false;
		}
	}

	return true;
}

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
