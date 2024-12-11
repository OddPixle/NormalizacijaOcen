<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

// Function to normalize scores based on min-max scaling to range 1-10
function normalizeScores($scores, $min = 1, $max = 10) {
    $minScore = min($scores);
    $maxScore = max($scores);
    if ($maxScore == $minScore) {
        return array_fill(0, count($scores), ($max + $min) / 2);
    }
    $normalized = [];
    foreach ($scores as $score) {
        $normalized[] = $min + (($score - $minScore) / ($maxScore - $minScore)) * ($max - $min);
    }
    return $normalized;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFiles'])) {
    $files = $_FILES['excelFiles'];
    $processedFiles = [];
    $tempFolder = "uploads/";

    // Ensure temp folder exists
    if (!file_exists($tempFolder)) {
        mkdir($tempFolder, 0777, true);
    }

    // Loop through uploaded files
    for ($i = 0; $i < count($files['name']); $i++) {
        $filePath = $tempFolder . basename($files['name'][$i]);
        move_uploaded_file($files['tmp_name'][$i], $filePath);

        try {
            // Load the Excel file
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getActiveSheet();

            // Define the static range (Rows 2-10, Columns B-E)
            $startRow = 11;
            $endRow = 17;
            $startColumn = 'C';
            $endColumn = 'I';

            // Process the specified range
            for ($rowIndex = $startRow; $rowIndex <= $endRow; $rowIndex++) {
                $scores = [];
                foreach (range($startColumn, $endColumn) as $col) {
                    $cellValue = $sheet->getCell("$col$rowIndex")->getValue();
                    if (is_numeric($cellValue)) {
                        $scores[] = $cellValue;
                    }
                }
                if (count($scores) > 0) {
                    $normalized = normalizeScores($scores, 1, 10);
                    foreach (range($startColumn, $endColumn) as $index => $col) {
                        if (isset($normalized[$index])) {
                            $sheet->setCellValue("$col$rowIndex", round($normalized[$index], 2));
                        }
                    }
                }
            }

            // Save the updated file
            $processedFileName = $tempFolder . 'processed_' . basename($files['name'][$i]);
            $writer = new Xlsx($spreadsheet);
            $writer->save($processedFileName);
            $processedFiles[] = $processedFileName;
        } catch (Exception $e) {
            echo "Error processing file " . htmlspecialchars($files['name'][$i]) . ": " . $e->getMessage() . "<br>";
        }
    }

    // Zip all processed files
    $zipFileName = "processed_files.zip";
    $zip = new ZipArchive();
    if ($zip->open($zipFileName, ZipArchive::CREATE | ZipArchive::OVERWRITE) === TRUE) {
        foreach ($processedFiles as $processedFile) {
            if (file_exists($processedFile)) {
                $zip->addFile($processedFile, basename($processedFile));
            }
        }
        $zip->close();

        // Serve the ZIP file for download
        header('Content-Type: application/zip');
        header('Content-Disposition: attachment; filename="' . basename($zipFileName) . '"');
        header('Content-Length: ' . filesize($zipFileName));
        readfile($zipFileName);

         // Cleanup
         foreach ($processedFiles as $processedFile) {
            unlink($processedFile);
        }
        unlink($zipFileName);
        exit;
    } else {
        echo "Error creating ZIP file.";
    }
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload and Process Excel Files</title>
</head>
<body>
    <h1>Upload and Process Excel Files</h1>
    <form action="" method="post" enctype="multipart/form-data">
        <label for="excelFiles">Select Excel files:</label>
        <input type="file" name="excelFiles[]" id="excelFiles" multiple>
        <button type="submit">Upload and Process</button>
    </form>
</body>
</html>
