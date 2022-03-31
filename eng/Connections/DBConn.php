<?php
# FileName="Connection_php_mysql.htm"
# Type="MYSQL"
# HTTP="true"
$hostname_DBConn = "localhost";
$database_DBConn = "sieraddb";
$username_DBConn = "tinton";
$password_DBConn = "hearts";
$DBConn = mysql_pconnect($hostname_DBConn, $username_DBConn, $password_DBConn) or die(mysql_error());
?>