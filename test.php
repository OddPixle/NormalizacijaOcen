<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;


function normalizeScores($scores) {
    $normalized = [];
    $avrage = array_sum($scores) / count($scores);
    echo "<br>Avrage: $avrage<br>";
    $standardDeviation=calculateStandardDeviation($scores);
    echo "<br>Standard deviation: $standardDeviation <br>";
    $minScore=min($scores);
    $maxScore=max($scores);
    foreach ($scores as $score) {
        //$normalized[] = ($score-$avrage)/$standardDeviation;
        $normalized[]=($score - $minScore) /($maxScore-$minScore);
    }
    $adapted = [];
    foreach ($normalized as $normal){
        $adapted[] = 5.5 + $normal * 4.5; 
    }
    echo "<br>Normalized<br>";
    return $adapted;
}

function calculateStandardDeviation($scores) {
    // Izračun povprečja
    $mean = array_sum($scores) / count($scores);

    // Izračun vsote kvadratov razlik od povprečja
    $sumSquaredDifferences = 0;
    foreach ($scores as $value) {
        $sumSquaredDifferences += pow($value - $mean, 2);
    }

    // Izračun standardnega odklona
    $variance = $sumSquaredDifferences / count($scores);
    $standardDeviation = sqrt($variance);

    return $standardDeviation;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFiles'])) {
    $files = $_FILES['excelFiles'];
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

            // Define the static range
            $startRow = 11;
            $endRow = 17;
            $startColumn = 'C';
            $endColumn = 'I';

            // Process the specified range
            for ($rowIndex = $startRow; $rowIndex <= $endRow; $rowIndex++) {
                foreach (range($startColumn, $endColumn) as $col) {
                    $cellValue = $sheet->getCell("$col$rowIndex")->getValue();
                    if (is_numeric($cellValue)) {
                        echo "Row $rowIndex, Column $col: $cellValue<br>";
                    }
                }
            }
            $scores = [];

            for ($rowIndex = $startRow; $rowIndex <= $endRow; $rowIndex++) {
                foreach (range($startColumn, $endColumn) as $col) {
                    $cellValue = $sheet->getCell("$col$rowIndex")->getValue();
                    if (is_numeric($cellValue)) {
                        $scores[] = $cellValue;
                    }
                }
            
            }
            if (count($scores) > 0) {
                // Normalize the scores without writing them back to the sheet
                $normalized = normalizeScores($scores);
                echo "<br>"; // Separate rows for clarity
                echo "Row $rowIndex: <br>";
                foreach ($scores as $index => $original) {
                    $normalizedValue = isset($normalized[$index]) ? round($normalized[$index], 2) : "N/A";
                    echo "Original: $original, Normalized: $normalizedValue<br>";
                }
            }
            
        } catch (Exception $e) {
            echo "Error processing file " . htmlspecialchars($files['name'][$i]) . ": " . $e->getMessage() . "<br>";
        }

        // Optionally, delete the uploaded file after testing
        unlink($filePath);
    }
    exit;
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Test Excel File Inputs</title>
</head>
<body>
    <h1>Test Excel File Inputs</h1>
    <form action="" method="post" enctype="multipart/form-data">
        <label for="excelFiles">Select Excel files:</label>
        <input type="file" name="excelFiles[]" id="excelFiles" multiple>
        <button type="submit">Upload and Test</button>
    </form>
</body>
</html>
