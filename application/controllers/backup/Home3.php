<?php
class Home extends CI_Controller {

public function index() {
    $file = $_FILES['excel']['tmp_name'];

    // Step 1: Extract the XLSX file manually
    $unzipped_folder = 'application/uploads/extracted/';
    $this->unzipXLSX($file, $unzipped_folder);

    // Step 2: Read and parse the XML manually
    $sheetData = $this->readFileContents($unzipped_folder . 'xl/worksheets/sheet1.xml');
    $sharedStringsData = $this->readFileContents($unzipped_folder . 'xl/sharedStrings.xml');

    // Parse sharedStrings.xml
    $sharedStrings = $this->parseSharedStrings($sharedStringsData);

    // Parse sheet1.xml
    $data = $this->parseSheetData($sheetData, $sharedStrings);

    // Insert data into the database
    if (!empty($data)) {
        $this->Model_home->insert_batch($data);
        $this->load->view('home', array('success' => 'Data has been successfully uploaded.'));
    } else {
        $this->load->view('home', array('error' => 'No valid data found to insert.'));
    }
}

private function unzipXLSX($file, $destination) {
    $fp = fopen($file, 'r');
    $zip = fread($fp, filesize($file));
    fclose($fp);

    $offset = 0;
    $index = 0;

    // Find the central directory
    while ($pos = strpos($zip, "PK\x05\x06", $offset)) {
        $offset = $pos + 22;
        $index++;
        if ($index > 10) break;
    }

    // Read the central directory
    if ($pos !== false) {
        $endDir = substr($zip, $pos, 22);
        $fileCount = unpack('v', substr($endDir, 10, 2))[1];
        $dirSize = unpack('V', substr($endDir, 12, 4))[1];
        $dirOffset = unpack('V', substr($endDir, 16, 4))[1];
        $dir = substr($zip, $dirOffset, $dirSize);

        // Extract files
        $offset = 0;
        for ($i = 0; $i < $fileCount; $i++) {
            $header = substr($dir, $offset, 46);
            $offset += 46;
            $fileNameLength = unpack('v', substr($header, 28, 2))[1];
            $extraFieldLength = unpack('v', substr($header, 30, 2))[1];
            $fileCommentLength = unpack('v', substr($header, 32, 2))[1];
            $localOffset = unpack('V', substr($header, 42, 4))[1];

            $fileName = substr($dir, $offset, $fileNameLength);
            $offset += $fileNameLength + $extraFieldLength + $fileCommentLength;

            // Read the file's content from the ZIP archive
            $localFileHeader = substr($zip, $localOffset, 30);
            $compressedSize = unpack('V', substr($localFileHeader, 18, 4))[1];
            $dataOffset = $localOffset + 30 + $fileNameLength + $extraFieldLength;
            $fileData = substr($zip, $dataOffset, $compressedSize);

            // Create directories if necessary
            $filePath = $destination . $fileName;
            $dirPath = dirname($filePath);
            if (!is_dir($dirPath)) {
                mkdir($dirPath, 0777, true);
            }

            // Write the extracted file
            file_put_contents($filePath, $fileData);
        }
    }
}

private function readFileContents($filePath) {
    return file_get_contents($filePath);
}

private function parseSharedStrings($xmlContent) {
    $strings = [];
    preg_match_all('/<si>(.*?)<\/si>/s', $xmlContent, $matches);

    foreach ($matches[1] as $match) {
        preg_match('/<t>(.*?)<\/t>/s', $match, $value);
        $strings[] = isset($value[1]) ? $value[1] : '';
    }

    return $strings;
}

private function parseSheetData($xmlContent, $sharedStrings) {
    $data = [];
    preg_match_all('/<row.*?>(.*?)<\/row>/s', $xmlContent, $rows);

    foreach ($rows[1] as $index => $row) {
        if ($index === 0) continue; // Skip header row

        preg_match_all('/<c.*?>(.*?)<\/c>/s', $row, $cells);
        $id = $this->getCellValue($cells[1][0], $sharedStrings);
        $nama = $this->getCellValue($cells[1][1], $sharedStrings);

        if (!empty($id) && !empty($nama)) {
            $data[] = array('id' => $id, 'nama' => $nama);
        }
    }

    return $data;
}

private function getCellValue($cell, $sharedStrings) {
    preg_match('/<v>(.*?)<\/v>/s', $cell, $value);

    if (strpos($cell, 't="s"') !== false) {
        // This is a shared string
        return isset($value[1]) ? $sharedStrings[(int)$value[1]] : '';
    } else {
        // This is a number or inline string
        return isset($value[1]) ? $value[1] : '';
    }
}
}
