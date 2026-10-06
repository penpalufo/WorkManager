<?php

const WORK_MANAGER_NAS_DIRECTORY = '\\\\192.168.100.4\\disk1\\WorkManager';
const WORK_MANAGER_EXCEL_PATH = WORK_MANAGER_NAS_DIRECTORY . '\\WorkManager.xlsx';
const WORK_MANAGER_LOCK_PATH = WORK_MANAGER_NAS_DIRECTORY . '\\WorkManager.lock';
const WORK_MANAGER_LOCK_TIMEOUT = 600;

function startWorkManagerSession()
{
	if (session_status() !== PHP_SESSION_ACTIVE) {
		session_start();
	}
}

function remoteAddress()
{
	return isset($_SERVER['REMOTE_ADDR']) ? (string) $_SERVER['REMOTE_ADDR'] : '';
}

function lockClientId()
{
	if (!empty($_SERVER['HTTP_X_WORKMANAGER_LOCK'])) {
		return (string) $_SERVER['HTTP_X_WORKMANAGER_LOCK'];
	}
	return isset($_GET['client']) ? (string) $_GET['client'] : '';
}

function sessionOwnsLock($clientId = null)
{
	if (empty($_SESSION['work_manager_lock'])) {
		return false;
	}
	if ($clientId !== null && (!isset($_SESSION['work_manager_lock_client']) ||
		!hash_equals((string) $_SESSION['work_manager_lock_client'], $clientId))) {
		return false;
	}

	$contents = @file_get_contents(WORK_MANAGER_LOCK_PATH);
	return $contents !== false && trim($contents) === remoteAddress();
}

function refreshOwnedLock($clientId = null)
{
	if (!sessionOwnsLock($clientId)) {
		return false;
	}

	$address = remoteAddress();
	$written = @file_put_contents(WORK_MANAGER_LOCK_PATH, $address, LOCK_EX);
	return $written === strlen($address);
}

function lockIsStale()
{
	clearstatcache(true, WORK_MANAGER_LOCK_PATH);
	$modifiedAt = @filemtime(WORK_MANAGER_LOCK_PATH);
	return $modifiedAt !== false && time() - $modifiedAt >= WORK_MANAGER_LOCK_TIMEOUT;
}

