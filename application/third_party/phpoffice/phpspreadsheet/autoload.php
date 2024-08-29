<?php
spl_autoload_register(function ($class) {
    // Define the base directory for PHPSpreadsheet namespace
    $baseDir = __DIR__ . '/src/';

    // Check if the class uses the PhpOffice\PhpSpreadsheet namespace
    $prefix = 'PhpOffice\\PhpSpreadsheet\\';
    $len = strlen($prefix);

    if (strncmp($prefix, $class, $len) !== 0) {
        // Not the PhpSpreadsheet namespace, move to the next autoloader
        return;
    }

    // Get the relative class name
    $relativeClass = substr($class, $len);

    // Replace namespace separators with directory separators
    $file = $baseDir . str_replace('\\', '/', $relativeClass) . '.php';

    // Include the file if it exists
    if (file_exists($file)) {
        require $file;
    }
});
