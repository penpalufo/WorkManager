<?php
require_once __DIR__ . '/lock-common.php';

header('Content-Type: application/json; charset=UTF-8');
header('Cache-Control: no-store, no-cache, must-revalidate');

startWorkManagerSession();
$action = isset($_GET['action']) ? $_GET['action'] : '';
$clientId = lockClientId();

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
	respond(405, false, false, 'POSTメソッドで送信してください。');
}

if ($action === 'acquire') {
	acquireLock();
}

if ($action === 'refresh') {
	if (!$clientId || !refreshOwnedLock($clientId)) {
		respond(409, false, true, 'ファイルロック中です。編集はできません。');
	}
	respond(200, true, false, 'ファイルロックを更新しました。');
}

if ($action === 'release') {
	if ($clientId && sessionOwnsLock($clientId)) {
		@unlink(WORK_MANAGER_LOCK_PATH);
		unset($_SESSION['work_manager_lock']);
		unset($_SESSION['work_manager_lock_client']);
	}
	respond(200, true, false, 'ファイルロックを解除しました。');
}

respond(400, false, false, '不明な操作です。');

function acquireLock()
{
	$clientId = lockClientId();
	if (!$clientId) {
		respond(400, false, false, 'ロック識別子がありません。');
	}

	if (sessionOwnsLock() && refreshOwnedLock()) {
		$_SESSION['work_manager_lock_client'] = $clientId;
		respond(200, true, false, 'ファイルロックを取得しました。');
	}

	unset($_SESSION['work_manager_lock']);
	if (file_exists(WORK_MANAGER_LOCK_PATH) && lockIsStale()) {
		@unlink(WORK_MANAGER_LOCK_PATH);
	}

	$handle = @fopen(WORK_MANAGER_LOCK_PATH, 'x');
	if ($handle === false) {
		respond(409, false, true, 'ファイルロック中です。編集はできません。');
	}

	$address = remoteAddress();
	$written = @fwrite($handle, $address);
	@fflush($handle);
	@fclose($handle);

	if ($written !== strlen($address)) {
		@unlink(WORK_MANAGER_LOCK_PATH);
		respond(500, false, false, 'ロックファイルを作成できませんでした。');
	}

	$_SESSION['work_manager_lock'] = true;
	$_SESSION['work_manager_lock_client'] = $clientId;
	respond(200, true, false, 'ファイルロックを取得しました。');
}

function respond($statusCode, $success, $locked, $message)
{
	http_response_code($statusCode);
	echo json_encode(array(
		'success' => $success,
		'locked' => $locked,
		'message' => $message
	), JSON_UNESCAPED_UNICODE);
	exit;
}

