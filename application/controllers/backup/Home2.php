<?php
defined('BASEPATH') OR exit('No direct script access allowed');

//require_once 'application/libraries/simplexlsx-master/src/SimpleXLSX.php'; // Pastikan SimpleXLSX ada di folder third_party
require_once APPPATH . 'libraries/simplexlsx-master/src/SimpleXLSX.php';

use Shuchkin\SimpleXLSX;

class Home extends CI_Controller {

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

            // Load the spreadsheet using SimpleXLSX
            if ($xlsx = SimpleXLSX::parse($inputFileName)) {
                $sheetData = $xlsx->rows();

                // Skip the header row and prepare the data to be inserted
                $data = array();
                $headerSkipped = false;
                foreach ($sheetData as $rowIndex => $row) {
                    if (!$headerSkipped) {
                        $headerSkipped = true; // Skip the first row (headers)
                        continue;
                    }

                    if (!empty($row[0]) && !empty($row[1])) {
                        $data[] = array(
                            'id' => $row[0],
                            'nama' => $row[1]
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
            } else {
                $this->load->view('home', array('error' => 'Error loading file: ' . SimpleXLSX::parseError()));
            }
        }
    }
}