<?php
// Load PhpSpreadsheet
require_once APPPATH . 'libraries/PhpSpreadsheet/src/Bootstrap.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Phpspreadsheet_lib {
    public function load($filePath) {
        return IOFactory::load($filePath);
    }
}