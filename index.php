<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//use ZipArchive;

function normalizeScores($scores) {
    $minScore = min($scores);
    $maxScore = max($scores);
    $normalized = array_map(function ($score) use ($minScore, $maxScore) {
        return 5.5 + (($score - $minScore) / ($maxScore - $minScore)) * 4.5;
    }, $scores);
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

    for ($i = 0; $i < count($files['name']); $i++) {
        $originalFilePath = $tempFolder . basename($files['name'][$i]);
        move_uploaded_file($files['tmp_name'][$i], $originalFilePath);

        try {
            $spreadsheet = IOFactory::load($originalFilePath);
            $sheet = $spreadsheet->getActiveSheet();

            $startRow = 11;
            $endRow = 17;
            $startColumn = 'C';
            $endColumn = 'I';

            $scores = [];

            // Define columns to skip
            $skipColumns = ['D', 'F', 'H', 'J'];

            foreach (range($startColumn, $endColumn) as $col) {
                if (in_array($col, $skipColumns)) {
                    continue; // Skip the specified columns
                }
                for ($rowIndex = $startRow; $rowIndex <= $endRow; $rowIndex++) {
                    $value = $sheet->getCell("$col$rowIndex")->getValue();
                    if (is_numeric($value)) {
                        $scores[] = $value;
                    }
                }
            }


            if (!empty($scores)) {
                $normalized = normalizeScores($scores);
                $index = 0;
                foreach (range($startColumn, $endColumn) as $col) {
                    for ($rowIndex = $startRow; $rowIndex <= $endRow; $rowIndex++) {
                        if (in_array($col, $skipColumns)) {
                            continue; // Skip the specified columns
                        }
                        if (isset($normalized[$index])) {
                            $sheet->setCellValue("$col$rowIndex", round($normalized[$index], 1));

                            // Apply number format with a comma as the decimal separator
                            $sheet->getStyle("$col$rowIndex")
                                ->getNumberFormat()
                                ->setFormatCode('#,##0.0');
                            $index++;
                        }
                    }
                }
            }

            $processedFilePath = $tempFolder . 'processed_' . basename($files['name'][$i]);
            $writer = new Xlsx($spreadsheet);
            $writer->save($processedFilePath);

            if (file_exists($processedFilePath)) {
                $processedFiles[] = $processedFilePath;
            } else {
                throw new Exception("Processed file $processedFilePath could not be created.");
            }
        } catch (Exception $e) {
            echo "Error processing file " . htmlspecialchars($files['name'][$i]) . ": " . $e->getMessage() . "<br>";
        }
    }

    $zipFileName = $tempFolder . "processed_files.zip";
    $zip = new ZipArchive();

    // Open the ZIP archive for creation
    if ($zip->open($zipFileName, ZipArchive::CREATE | ZipArchive::OVERWRITE) === TRUE) {
        foreach ($processedFiles as $processedFile) {
            if (file_exists($processedFile)) {
                $zip->addFile($processedFile, basename($processedFile));
            } else {
                echo "Warning: $processedFile does not exist.<br>";
            }
        }
        $zip->close();
    } else {
        echo "Error: Unable to create ZIP file.<br>";
        exit;
    }

    // Serve the ZIP file for download
    if (file_exists($zipFileName)) {
        header('Content-Type: application/zip');
        header('Content-Disposition: attachment; filename="' . basename($zipFileName) . '"');
        header('Content-Length: ' . filesize($zipFileName));

        // Clean buffer to avoid corruption
        ob_clean();
        flush();
        readfile($zipFileName);

        // Clean up files after download
        foreach ($processedFiles as $processedFile) {
            unlink($processedFile);
        }
        unlink($zipFileName);
        exit;
    } else {
        echo "Error: ZIP file could not be found.<br>";
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
