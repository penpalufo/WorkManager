<?php
// ブラウザで変換・重複除去したマスターを、固定のローカルパスへ保存する
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
$master = json_decode(file_get_contents('php://input'), true);
if (!is_array($master) || !isset($master['sheetName'], $master['rows'])
    || $master['sheetName'] !== 'customer' || !is_array($master['rows'])) {
    respond(400, false, 'マスターのJSON形式が不正です。');
}
// rowsは見出しをキーにしたオブジェクト配列として受け取る
$objectMaster = json_decode(json_encode($master));
foreach ($objectMaster->rows as $row) {
    if (!is_object($row)) {
        respond(400, false, 'rowsには取引先オブジェクトを指定してください。');
    }
}
$master = array('sheetName' => 'customer', 'rows' => $objectMaster->rows);
$directory = dirname(__DIR__) . '/data';
if (!is_dir($directory) && !@mkdir($directory, 0775, true) && !is_dir($directory)) {
    respond(500, false, 'dataディレクトリを作成できません。');
}
$json = json_encode($master, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
if ($json === false) {
    respond(400, false, 'JSONへ変換できません。');
}
// 書き込みが完了したファイルだけを公開する
$temp = @tempnam($directory, 'customer-');
if ($temp === false) {
    respond(500, false, 'dataディレクトリに書き込めません。');
}
if (@file_put_contents($temp, $json, LOCK_EX) !== strlen($json)) {
    @unlink($temp);
    respond(500, false, 'JSONの書き込みに失敗しました。');
}
if (!@rename($temp, $directory . '/customerMaster.json')) {
    @unlink($temp);
    respond(500, false, 'customerMaster.jsonの保存に失敗しました。');
}
respond(200, true, '取引先名マスターをJSONへ保存しました。');
