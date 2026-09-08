<?php
// 最大約1MBずつ受信し、全件をメモリに展開せず追記する
header('Content-Type: application/json; charset=UTF-8');
header('Cache-Control: no-store');
function respond($status, $message, $offset = 0) {
    http_response_code($status);
    echo json_encode(array('success' => $status === 200, 'message' => $message, 'offset' => $offset), JSON_UNESCAPED_UNICODE);
    exit;
}
if ($_SERVER['REQUEST_METHOD'] !== 'POST') { header('Allow: POST'); respond(405, 'POSTで送信してください。'); }
$input = file_get_contents('php://input', false, null, 0, 1100001);
if ($input === false || strlen($input) > 1100000) respond(413, '送信サイズが大きすぎます。');
$batch = json_decode($input);
if (!is_object($batch) || !isset($batch->id, $batch->offset, $batch->final, $batch->rows)
    || !is_string($batch->id) || !preg_match('/^[a-f0-9]{32}$/D', $batch->id)
    || !is_int($batch->offset) || $batch->offset < 0 || !is_bool($batch->final)
    || !is_array($batch->rows) || count($batch->rows) > 500) respond(400, '分割データの形式が不正です。');
$parts = array();
foreach ($batch->rows as $row) {
    if (!is_object($row)) respond(400, '担当者データが不正です。');
    $parts[] = json_encode($row, JSON_UNESCAPED_UNICODE);
}
$directory = dirname(__DIR__) . '/data';
if (!is_dir($directory) && !@mkdir($directory, 0775, true) && !is_dir($directory)) respond(500, 'dataを作成できません。');
$path = $directory . '/contact-' . $batch->id . '.tmp';
$file = @fopen($path, $batch->offset === 0 ? 'x+b' : 'r+b');
if (!$file) respond(409, '一時ファイルを開けません。再読み込みしてください。');
if (!flock($file, LOCK_EX)) { fclose($file); respond(500, 'ロックできません。'); }
$size = fstat($file)['size'];
if ($size !== $batch->offset) { fclose($file); respond(409, '送信順序が一致しません。'); }
$text = $size === 0 ? '{"sheetName":"customerContact","rows":[' : '';
if ($parts) $text .= ($size > 0 ? ',' : '') . implode(',', $parts);
if ($batch->final) $text .= ']}';
fseek($file, 0, SEEK_END);
$written = 0;
while ($written < strlen($text)) {
    $n = fwrite($file, substr($text, $written));
    if ($n === false || $n === 0) { ftruncate($file, $size); fclose($file); respond(500, '書き込みに失敗しました。'); }
    $written += $n;
}
$ok = fflush($file);
fclose($file);
if (!$ok) respond(500, '書き込みを確定できません。');
// 完了するまでは既存のJSONを変更しない
if ($batch->final && !@rename($path, $directory . '/customerContactMaster.json')) respond(500, 'JSONの公開に失敗しました。');
respond(200, '受信しました。', $size + $written);
