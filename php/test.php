<?php
echo "<h3>サーバー変数の一覧</h3>";
echo "REMOTE_ADDR: " . ($_SERVER['REMOTE_ADDR'] ?? 'なし') . "<br>";
echo "HTTP_X_FORWARDED_FOR: " . ($_SERVER['HTTP_X_FORWARDED_FOR'] ?? 'なし') . "<br>";
echo "HTTP_X_REAL_IP: " . ($_SERVER['HTTP_X_REAL_IP'] ?? 'なし') . "<br>";

// デバッグ用：すべてのサーバー情報を出力して探す
echo "<h4>すべてのサーバー変数（Ctrl + F で 192.168. を探してみてください）</h4>";
echo "<pre>";
print_r($_SERVER);
echo "</pre>";
