<?php

// Load PhpSpreadsheet
require_once APPPATH . 'libraries/simplexlsx-master/src/SimpleXLSX.php';

use Shuchkin\SimpleXLSX;
//use Shuchkin\SimpleXLSX\SimpleXLSX;

class Home extends CI_Controller {

    public function __construct() {
        parent::__construct();
        $this->load->model('Model_home');
        $this->load->library('upload'); // Load upload library
    }

    public function index() {
        // Konfigurasi upload file
        $config['upload_path'] = 'application/uploads/';
        $config['allowed_types'] = 'xlsx|csv';
        $config['max_size'] = 10000;

        // Inisialisasi library upload
        $this->upload->initialize($config);

        if (!$this->upload->do_upload('excel')) {
            $error = array('error' => $this->upload->display_errors());
            $this->load->view('home', $error);
        } else {
            $file = $this->upload->data();
            $inputFileName = $file['full_path'];
            $fileExtension = strtolower(pathinfo($inputFileName, PATHINFO_EXTENSION));

            $data = array();

            if ($fileExtension === 'csv') {
                // Handle CSV
                if (($handle = fopen($inputFileName, 'r')) !== FALSE) {
                    $headerSkipped = false;
                    while (($row = fgetcsv($handle, 1000, ',')) !== FALSE) {
                        if (!$headerSkipped) {
                            $headerSkipped = true;
                            continue;
                        }
                        if (!empty($row[0]) && !empty($row[1])) {
                            $data[] = array(
                                'adjustmenttype' => $row[0], 
                                'day' => $row[1],
                                'country' => $row[2],
                                'videoid' => $row[3],
                                'videotitle' => $row[4],
                                'videoduration' => $row[5],
                                'category' => $row[6],
                                'channelid' => $row[7],
                                'uploader' => $row[8],
                                'channeldisplayname' => $row[9],
                                'contenttype' => $row[10],
                                'policy' => $row[11],
                                'ownedviews' => $row[12],
                                'YouTube Revenue Split : Auction' => $row[13],
                                'YouTube Revenue Split : Reserved' => $row[14],
                                'YouTube Revenue Split : Partner Sold YouTube Served' => $row[15],
                                'YouTube Revenue Split : Partner Sold Partner Served' => $row[16],
                                'YouTube Revenue Split' => $row[17],
                                'Partner Revenue : Auction' => $row[18],
                                'Partner Revenue : Reserved' => $row[19],
                                'Partner Revenue : Partner Sold YouTube Served' => $row[20],
                                'Partner Revenue : Partner Sold Partner Served' => $row[21],
                                'Partner Revenue' => $row[22],
                            );
                        }
                    }
                    fclose($handle);
                } else {
                    $this->load->view('home', array('error' => 'Error opening CSV file.'));
                    return;
                }
            } elseif ($fileExtension === 'xlsx') {
                // Handle XLSX using SimpleXLSX
                if ($xlsx = SimpleXLSX::parse($inputFileName)) {
                    $headerSkipped = false;

                    foreach ($xlsx->rows() as $row) {
                        if (!$headerSkipped) {
                            $headerSkipped = true;
                            continue;
                        }
                        if (!empty($row[0]) && !empty($row[1])) {
                            $data[] = array('id' => $row[0], 'nama' => $row[1]);
                        }
                    }
                } else {
                    $this->load->view('home', array('error' => 'Error reading XLSX file.'));
                    return;
                }
            }

            // Insert data ke database
            if (!empty($data)) {
                $this->Model_home->insert_batch($data);
                $this->load->view('home', array('success' => 'Data has been successfully uploaded.'));
            } else {
                $this->load->view('home', array('error' => 'No valid data found to insert.'));
            }
        }
    }
}