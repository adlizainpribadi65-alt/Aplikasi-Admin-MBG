<?php
// ============================================================
//  API USERS SYNC — ADMIN MBG
//  Simpan dan ambil akun registrasi agar bisa diakses antar device
// ============================================================

header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit;
}

$dataFile = __DIR__ . '/users_db.json';

// Default users jika file belum ada
$defaultUsers = [
    [
        'username' => 'admin',
        'password' => 'admin123',
        'name'     => 'Adli Zain Pribadi',
        'role'     => 'Administrator'
    ],
    [
        'username' => 'operator',
        'password' => 'op123',
        'name'     => 'Operator SPPG',
        'role'     => 'Operator'
    ],
    [
        'username' => 'viewer',
        'password' => 'view123',
        'name'     => 'Pengawas Audit',
        'role'     => 'Viewer'
    ]
];

function readUsers($file, $defaults) {
    if (!file_exists($file)) {
        file_put_contents($file, json_encode($defaults, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));
        return $defaults;
    }
    $raw = @file_get_contents($file);
    if (!$raw) return $defaults;
    $parsed = json_decode($raw, true);
    return is_array($parsed) ? $parsed : $defaults;
}

function writeUsers($file, $users) {
    return file_put_contents($file, json_encode($users, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE), LOCK_EX);
}

$method = $_SERVER['REQUEST_METHOD'];

if ($method === 'GET') {
    $users = readUsers($dataFile, $defaultUsers);
    echo json_encode(['status' => 'success', 'data' => $users]);
    exit;
}

if ($method === 'POST') {
    $input = file_get_contents('php://input');
    $body  = json_decode($input, true);

    if (!$body || !isset($body['username']) || !isset($body['password'])) {
        http_response_code(400);
        echo json_encode(['status' => 'error', 'message' => 'Data tidak lengkap']);
        exit;
    }

    $users = readUsers($dataFile, $defaultUsers);

    // Cek duplikat username (case-insensitive)
    $usernameLower = strtolower(trim($body['username']));
    foreach ($users as $u) {
        if (strtolower($u['username']) === $usernameLower) {
            http_response_code(409);
            echo json_encode(['status' => 'error', 'message' => 'Username sudah terdaftar']);
            exit;
        }
    }

    $newUser = [
        'username'  => trim($body['username']),
        'password'  => $body['password'],
        'name'      => isset($body['name']) ? trim($body['name']) : trim($body['username']),
        'role'      => isset($body['role']) ? trim($body['role']) : 'Operator',
        'createdAt' => date('Y-m-d H:i:s')
    ];

    $users[] = $newUser;
    writeUsers($dataFile, $users);

    echo json_encode([
        'status'  => 'success',
        'message' => 'Akun berhasil didaftarkan di server dan dapat diakses dari device lain',
        'data'    => $newUser
    ]);
    exit;
}

http_response_code(405);
echo json_encode(['status' => 'error', 'message' => 'Metode tidak didukung']);
