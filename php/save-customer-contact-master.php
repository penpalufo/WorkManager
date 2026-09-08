<?php
// 変換済みの担当者マスターを固定のローカルパスへ保存する
header('Content-Type: application/json; charset=UTF-8');
header('Cache-Control: no-store');

function respond($status, $success, $message)
{
    http_response_code($status);
    echo json_encode(array('success' => $success, 'message' => $message), JSON_UNESCAPED_UNICODE);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Allow: POST');
    respond(405, false, 'POSTメソッドで送信してください。');
}
$master = json_decode(file_get_contents('php://input'));
if (!is_object($master) || !isset($master->sheetName, $master->rows)
    || $master->sheetName !== 'customerContact' || !is_array($master->rows)) {
    respond(400, false, '担当者マスターのJSON形式が不正です。');
}
foreach ($master->rows as $row) {
    if (!is_object($row)) {
        respond(400, false, 'rowsには担当者オブジェクトを指定してください。');
    }
}
$directory = dirname(__DIR__) . '/data';
if (!is_dir($directory) && !@mkdir($directory, 0775, true) && !is_dir($directory)) {
    respond(500, false, 'dataディレクトリを作成できません。');
}
$json = json_encode(array('sheetName' => 'customerContact', 'rows' => $master->rows),
    JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
if ($json === false) {
    respond(400, false, 'JSONへ変換できません。');
}
// 一時ファイルへの全量書き込みを確認してから公開する
$temp = @tempnam($directory, 'contact-');
if ($temp === false) {
    respond(500, false, 'dataディレクトリに書き込めません。');
}
if (@file_put_contents($temp, $json, LOCK_EX) !== strlen($json)) {
    @unlink($temp);
    respond(500, false, 'JSONの書き込みに失敗しました。');
}
if (!@rename($temp, $directory . '/customerContactMaster.json')) {
    @unlink($temp);
    respond(500, false, 'customerContactMaster.jsonの保存に失敗しました。');
}
respond(200, true, '取引先担当者マスターをJSONへ保存しました。');
