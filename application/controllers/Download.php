<?php
defined('BASEPATH') OR exit('No direct script access allowed');

//Library PhpSpreadsheet //Berhasil tapi pakai composer
require('vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


//Library BoxSpout //Gagal
require('application/libraries/spout-3.3.0/spout-3.3.0/src/Spout/Autoloader/autoload.php');
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Common\Entity\Row;

//Library SimpleXLSXGen //Berhasil dilocal tapi gagal di live
require('application/libraries/simplexlsxgen-master/simplexlsxgen-master/src/SimpleXLSXGen.php');
use Shuchkin\SimpleXLSXGen;

//Library XLSXWriter //Gagal
require('application/libraries/PHP_XLSXWriter-master/PHP_XLSXWriter-master/XLSXWriter.class.php');
//require('application/libraries/PHP_XLSXWriter-master/PHP_XLSXWriter-master/testbench/XLSXWriter.class.Test.php');
//use XLSXWriter;

// Load the PHP Excel Writer library //Gagal
require('application/libraries/php-excel-writer-master/php-excel-writer-master/src/ExcelWriter.php');
use Ellumilel\ExcelWriter;


class Download extends CI_Controller {

	public function __construct() {
        parent::__construct();
        $this->load->model('Model_home');
        //$this->load->library('upload'); // Load upload library
    }

	public function index()
	{
		$this->load->view('download');
	}

	/*
	//Library Spreadsheet //Berhasil tapi pakai autoload
	public function excel(){
		$data['download'] = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();

		//var_dump($data['download']);
		//$this->Model_home->tampil_data($data)->result();

		//require(APPPATH. 'PHPExcel-1.8/Classes/PHPExcel.php');
		//require(APPPATH. 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php');

		//$object = new PHPExcel();
		$spreadsheet = new Spreadsheet();

		//$object->getProperties()->setCreator("Wilfredo Alexander Sutanto");
		//$object->getProperties()->setLastModifiedBy("Wilfredo Alexander Sutanto");
		//$object->getProperties()->setTitle("Daftar Pusatmusik");

		$spreadsheet->getProperties()->setCreator("Wilfredo Alexander Sutanto")
        ->setLastModifiedBy("Wilfredo Alexander Sutanto")
        ->setTitle("Daftar Pusatmusik");

		//$object->setActiveSheetIndex(0);
		$sheet = $spreadsheet->getActiveSheet();

		//$object->getActiveSheet()->setCellValue('A1', 'No');
		//$object->getActiveSheet()->setCellValue('B1', 'Title');
		//$object->getActiveSheet()->setCellValue('C1', 'Name');
		//$object->getActiveSheet()->setCellValue('D1', 'Country');
		//$object->getActiveSheet()->setCellValue('E1', 'Average View Duration');
		//$object->getActiveSheet()->setCellValue('F1', 'Views');

		$sheet->setCellValue('A1', 'No');
		$sheet->setCellValue('B1', 'Title');
		$sheet->setCellValue('C1', 'Name');
		$sheet->setCellValue('D1', 'Country');
		$sheet->setCellValue('E1', 'Average View Duration');
		$sheet->setCellValue('F1', 'Views');

		$baris = 2;
		$no = 1;

		foreach($data['download'] as $download) {

			//$object->getActiveSheet()->setCellValue('A'.$baris, $no++);
			//$object->getActiveSheet()->setCellValue('B'.$baris, $download->title);
			//$object->getActiveSheet()->setCellValue('C'.$baris, $download->name);
			//$object->getActiveSheet()->setCellValue('D'.$baris, $download->country);
			//$object->getActiveSheet()->setCellValue('E'.$baris, $download->averageviewduration);
			//$object->getActiveSheet()->setCellValue('F'.$baris, $download->views);
			
			//$sheet->setCellValue('A'.$baris, $no++);
			//$sheet->setCellValue('B'.$baris, $download->title);
			//$sheet->setCellValue('C'.$baris, $download->name);
			//$sheet->setCellValue('D'.$baris, $download->country);
			//$sheet->setCellValue('E'.$baris, $download->averageviewduration);
			//$sheet->setCellValue('F'.$baris, $download->views);

			$sheet->setCellValue('A'.$baris, $no++);
			$sheet->setCellValue('B'.$baris, "ABC");
			$sheet->setCellValue('C'.$baris, "ABC");
			$sheet->setCellValue('D'.$baris, "ABC");
			$sheet->setCellValue('E'.$baris, "ABC");
			$sheet->setCellValue('F'.$baris, "ABC");

			$baris++;
		}

		//$filename="Data_Pusatmusik".'xlsx';
		$filename="Data_Pusatmusik.xlsx";

		//$object->getActiveSheet->setTitle("Data Pusatmusik");
		$sheet->setTitle("Data Pusatmusik");

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="'.$filename.'"');
		header('Cache-Control: max-age=0');

		//$writer=PHPExcel_IOFactory::createwriter($object, 'Excel2007');
		$writer = new Xlsx($spreadsheet);

		ob_clean();
		flush();
		$writer->save('php://output');

		exit;
	}
	*/
	
	
	/*
	//Library Box Spout //Gagal
	public function excel(){
		// Mulai output buffering
		ob_start();

		// Ambil data mahasiswa
	    $data['download'] = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();

		// Bersihkan buffer output sebelum menulis file
		//ob_clean();
		//flush();

		// Membuat writer untuk Excel / XLSX
	    $writer = WriterEntityFactory::createXLSXWriter();
		$filename = 'Data_Pusatmusik.xlsx';

		// Mengatur header agar browser menangani file sebagai file Excel
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

	    $writer->openToBrowser($filename);

	    // Buat header
	    $headerRow = WriterEntityFactory::createRowFromArray([
	        'No', 'Title', 'Name', 'Country', 'Average View Duration', 'Views'
	    ]);
	    $writer->addRow($headerRow);

	    // Loop untuk setiap baris data mahasiswa
	    $no = 1;
	    foreach ($data['download'] as $download) {
	        $row = WriterEntityFactory::createRowFromArray([
	            $no++,
	            $download->title,
	            $download->name,
	            $download->country,
	            $download->averageviewduration,
	            $download->views,
	        ]);
	        $writer->addRow($row);
	    }

	    // Menutup setelah selesai menulis writer
	    $writer->close();

		// Bersihkan output buffer setelah file dikirim
		//ob_end_clean();
		exit;
	}
	*/

	/*
	//CSV //Berhasil tapi format CSV
	public function excel(){
		// Set nama file CSV
	    $filename = "Data_Pusatmusik.csv";

	    // Ambil data mahasiswa dari model
	    $data['download'] = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();

	    // Set header untuk mendownload file CSV
	    header('Content-Type: text/csv');
	    header('Content-Disposition: attachment;filename="'.$filename.'"');
	    header('Cache-Control: max-age=0');

	    // Buka output ke PHP output stream (php://output)
	    $output = fopen('php://output', 'w');

	    // Tulis header CSV
	    $header = ['No', 'Title', 'Name', 'Country', 'Average View Duration', 'Views'];
	    fputcsv($output, $header);

	    // Tulis data mahasiswa ke CSV
	    $no = 1;
	    foreach ($data['download'] as $download) {
	        $row = [
	            //$no++,
	            //$download->title,
	            //$download->name,
	            //$download->country,
	            //$download->averageviewduration,
	            //$download->views,

				$no++,
				"ABC",
				"ABC",
				"ABC",
				"ABC",
				"ABC",
	        ];
	        fputcsv($output, $row); // Menulis baris ke file CSV
	    }

	    // Tutup output
	    fclose($output);
	    exit;
	}
	*/

	/*
	//Simple XLSX Gen //Berhasil di local tapi gagal di live
	public function excel() {
        // Ambil data dari database
        $data = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();

        // Siapkan data untuk ditulis ke file Excel
        $header = ['No', 'Title', 'Name', 'Country', 'Average View Duration', 'Views'];
        $rows = [];
        $no = 1;

        foreach ($data as $download) {
            $rows[] = [
                $no++,
                $download->title,
                $download->name,
                $download->country,
                $download->averageviewduration,
                $download->views
            ];
        }

        // Gabungkan header dan rows
        $excelData = array_merge([$header], $rows);

        // Buat file Excel
        $xlsx = SimpleXLSXGen::fromArray($excelData);
        $xlsx->downloadAs('Data_Pusatmusik.xlsx'); // Download file

        exit;
    }
	*/

    /*
	//Php XLSX Writer //Gagal
	public function excel() {
		$data = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();
		
		$filename = "Data_Pusatmusik.xlsx";
		$header = ['No', 'Title', 'Name', 'Country', 'Average View Duration', 'Views'];
		$rows = [];
		$no = 1;

		foreach ($data as $download) {
			$rows[] = [
				$no++,
				$download->title,
				$download->name,
				$download->country,
				$download->averageviewduration,
				$download->views
			];
		}

		// Write to XLSX
		$writer = new XLSXWriter();
		$writer->writeSheetHeader('Sheet1', $header);

		foreach ($rows as $row) {
			$writer->writeSheetRow('Sheet1', $row);
		}

		$writer->writeToFile($filename);

		// Download the file
		header('Content-disposition: attachment; filename="'. $filename .'"');
		header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		header('Content-Length: '.filesize($filename));
		readfile($filename);
		unlink($filename); // Delete file after download

		exit;
	}
	*/

	/*
	//Php Excel Writer //Gagal
	public function excel() {
        // Load the PHP Excel Writer library
        //require('application/libraries/PHPExcelWriter.php');

        // Prepare the data
        $data = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();

        // Bersihkan buffer output sebelum menulis file
		//ob_clean();
		//flush();

        // Set the filename and header
        $filename = "Data_Pusatmusik.xls"; // Note: Using .xls for compatibility
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        // Create a new PHP Excel Writer object
        $excel = new PHPExcel();
        $excel->setActiveSheetIndex(0);
        $sheet = $excel->getActiveSheet();

        // Set the header row
        $sheet->setCellValue('A1', 'No');
        $sheet->setCellValue('B1', 'Title');
        $sheet->setCellValue('C1', 'Name');
        $sheet->setCellValue('D1', 'Country');
        $sheet->setCellValue('E1', 'Average View Duration');
        $sheet->setCellValue('F1', 'Views');

        // Fill data
        $no = 1;
        $rowNum = 2; // Starting from the second row
        foreach ($data as $download) {
            $sheet->setCellValue('A' . $rowNum, $no++);
            $sheet->setCellValue('B' . $rowNum, $download->title);
            $sheet->setCellValue('C' . $rowNum, $download->name);
            $sheet->setCellValue('D' . $rowNum, $download->country);
            $sheet->setCellValue('E' . $rowNum, $download->averageviewduration);
            $sheet->setCellValue('F' . $rowNum, $download->views);
            $rowNum++;
        }

        // Create the Excel file
        $writer = \PHPExcel_IOFactory::createWriter($excel, 'Excel5');
        ob_end_clean(); // Clean output buffer
        $writer->save('php://output'); // Output to browser

        exit; // Ensure no further output
    }
    */  

    //Format XLS //Berhasil dan digunakan sekarang
    public function excel() {
        // Get data from the database
        $data['download'] = $this->Model_home->tampil_data('tb_download_pusatmusik')->result();

        // Set the filename
        $filename = "Data_Pusatmusik.xls";

        // Set headers for the Excel file
        header("Content-Type: application/vnd.ms-excel");
        header("Content-Disposition: attachment; filename=\"$filename\"");
        header("Cache-Control: max-age=0");

        // Output the header row
        echo "No\tTitle\tName\tCountry\tAverage View Duration\tViews\n";

        // Output data rows
        $no = 1;
        foreach ($data['download'] as $download) {
            echo "$no\t{$download->title}\t{$download->name}\t{$download->country}\t{$download->averageviewduration}\t{$download->views}\n";
            $no++;
        }

        exit; // Ensure no further output
    }
}
