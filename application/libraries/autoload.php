<?php
// Fungsi autoload
spl_autoload_register(function ($class) {
    // Tentukan namespace dasar
    $namespace = 'PhpOffice\\PhpSpreadsheet\\';
    $baseDir = __DIR__ . '/PhpSpreadsheet/src/PhpSpreadsheet/';

    // Cek apakah kelas dimulai dengan namespace yang diinginkan
    if (strpos($class, $namespace) === 0) {
        // Hilangkan namespace dari nama kelas
        $relativeClass = substr($class, strlen($namespace));
        // Ganti namespace separator dengan separator direktori
        $file = $baseDir . str_replace('\\', '/', $relativeClass) . '.php';

        // Jika file ada, sertakan file
        if (file_exists($file)) {
            require_once $file;
        }
    }
});
