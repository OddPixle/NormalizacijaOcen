<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Function to calculate the average
function calculateAverage($values) {
    return array_sum($values) / count($values);
}

// Function to calculate the standard deviation
function calculateStandardDeviation($values, $average) {
    $sum = 0;
    foreach ($values as $value) {
        $sum += pow($value - $average, 2);
    }
    return sqrt($sum / count($values));
}

// Function to normalize scores
function normalizeScores($scores, $average, $stdDev) {
    $normalized = [];
    foreach ($scores as $score) {
        $normalized[] = ($stdDev != 0) ? ($score - $average) / $stdDev : 0;
    }
    return $normalized;
}

// Function to rescale scores back to the range 1–10
function rescaleScores($normalizedScores, $min = 1, $max = 10) {
    $rescaled = [];
    foreach ($normalizedScores as $score) {
        $rescaled[] = round(5 + $score * (($max - $min) / 2), 2);
    }
    return $rescaled;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFiles'])) {
    $files = $_FILES['excelFiles'];
    $processedData = [];
    $tempFolder = "uploads/";

    // Ensure temp folder exists
    if (!file_exists($tempFolder)) {
        mkdir($tempFolder, 0777, true);
    }

    // Loop through uploaded files
    for ($i = 0; $i < count($files['name']); $i++) {
        $filePath = $tempFolder . basename($files['name'][$i]);
        move_uploaded_file($files['tmp_name'][$i], $filePath);

        // Load the Excel file
        $spreadsheet = IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray();

        // Process data
        foreach ($data as $row) {
            if (empty($row)) continue; // Skip empty rows
            $scores = array_filter($row, 'is_numeric'); // Get numeric scores only
            if (count($scores) > 0) {
                $average = calculateAverage($scores);
                $stdDev = calculateStandardDeviation($scores, $average);
                $normalized = normalizeScores($scores, $average, $stdDev);
                $rescaled = rescaleScores($normalized);
                $processedData[] = $rescaled; // Add processed row
            }
        }

        // Optionally, delete the file after processing
        unlink($filePath);
    }

    // Create a new spreadsheet for the output
    $outputSpreadsheet = new Spreadsheet();
    $outputSheet = $outputSpreadsheet->getActiveSheet();

    // Write processed data to the new sheet
    foreach ($processedData as $rowIndex => $row) {
        foreach ($row as $colIndex => $value) {
            // Convert column index to column letter
            $columnLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex + 1);
            $outputSheet->setCellValue($columnLetter . ($rowIndex + 1), $value);
        }
    }
    

    // Output the new Excel file
    $outputFileName = "processed_data.xlsx";
    $writer = new Xlsx($outputSpreadsheet);
    $writer->save($outputFileName);

    // Offer the file for download
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . $outputFileName . '"');
    readfile($outputFileName);

    // Clean up the temporary output file
    unlink($outputFileName);
    exit;
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
