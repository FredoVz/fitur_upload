<?php
defined('BASEPATH') OR exit('No direct script access allowed');

//require_once 'vendor/autoload.php';
// Masukkan file PhpSpreadsheet secara manual

require_once('application/libraries/simple-cache-master/src/CacheInterface.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Collection/Memory/SimpleCache1.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/NumberFormat/BaseFormatter.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/IValueBinder.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/BaseParserClass.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/IComparable.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Supervisor.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/Dimension.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/IReadFilter.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/IOFactory.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Spreadsheet.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Shared/File.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Shared/Date.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/IReader.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/BaseReader.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/Namespaces.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/DefaultReadFilter.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/ReferenceHelper.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Security/XmlScanner.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Settings.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Calculation/Calculation.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Calculation/Category.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Calculation/Engine/CyclicReferenceStack.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Calculation/Engine/Logger.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Calculation/Engine/BranchPruner.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Theme.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/Worksheet.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Shared/StringHelper.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Collection/CellsFactory.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Collection/Cells.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Collection/Memory/SimpleCache3.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/PageSetup.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/PageMargins.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/HeaderFooter.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/SheetView.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/Protection.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/RowDimension.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/ColumnDimension.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/AutoFilter.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/AutoFilter/Column/Rule.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Document/Properties.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Shared/IntOrFloat.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Document/Security.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Style.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Font.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Color.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Fill.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Borders.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Border.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Alignment.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/NumberFormat.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/Protection.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/Styles.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/Theme.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/Properties.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/SheetViews.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Worksheet/Validations.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/AddressRange.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/Coordinate.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/SheetViewOptions.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/ColumnAndRowAttributes.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Calculation/Functions.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/Cell.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/DataType.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/IgnoredErrors.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/CellAddress.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Cell/DefaultValueBinder.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/PageSetup.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/Hyperlinks.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Reader/Xlsx/WorkbookView.php');
require_once('application/libraries/PhpSpreadsheet-master/src/PhpSpreadsheet/Style/NumberFormat/Formatter.php');

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Home1 extends CI_Controller {

    public function __construct() {
        parent::__construct();
    }

    public function index() {

        $this->load->model('Model_home');
        $this->load->library('upload'); // Load upload library

        // Configuration for file upload
        $config['upload_path'] = 'application/uploads/';
        $config['allowed_types'] = 'xls|xlsx';
        $config['max_size'] = 10000;

        // Initialize the upload library with configuration
        $this->upload->initialize($config);

        if (!$this->upload->do_upload('excel')) {
            $error = array('error' => $this->upload->display_errors());
            $this->load->view('home', $error);
        } else {
            $file = $this->upload->data();
            $inputFileName = $file['full_path'];

            // Load the spreadsheet using PHPSpreadsheet
            try {
                $spreadsheet = IOFactory::load($inputFileName);
                $sheet = $spreadsheet->getActiveSheet();
                $sheetData = $sheet->toArray(null, true, true, true);

                // Skip the header row and prepare the data to be inserted
                $data = array();
                $headerSkipped = false;
                foreach ($sheetData as $rowIndex => $row) {
                    if (!$headerSkipped) {
                        $headerSkipped = true; // Skip the first row (headers)
                        continue;
                    }

                    if (!empty($row['A']) && !empty($row['B'])) {
                        $data[] = array(
                            'id' => $row['A'],
                            'nama' => $row['B']
                        );
                    }
                }

                // Insert the data into the database
                if (!empty($data)) {
                    $this->Model_home->insert_batch($data);
                    $this->load->view('home', array('success' => 'Data has been successfully uploaded.'));
                } else {
                    $this->load->view('home', array('error' => 'No valid data found to insert.'));
                }
            } catch (Exception $e) {
                $this->load->view('home', array('error' => 'Error loading file: ' . $e->getMessage()));
            }
        }
    }
}