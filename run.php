<?php

namespace FashionDeals;

use PhpOffice\PhpSpreadsheet as Excel;
use Exception;

// Initialization
ini_set('memory_limit', '2G');
chdir(dirname(__FILE__));

// Validate input parameters
if ($argc !== 3) {
    echo 'Wrong execution parameters!' . PHP_EOL;
    exit(1);
}

// Check input and output files
$inFile  = $argv[1];
$outFile = $argv[2];

if (!is_file($inFile) || !is_readable($inFile)) {
    echo 'File "' . $inFile . '" does not exists or is not readable!';
    exit(1);
}

if (file_exists($outFile)) {
    echo 'Warning! File "' . $outFile . '" already exists!' . PHP_EOL
         .
         'If you will not break script execution (Ctrl+C / Cmd+C) within 10 seconds, it will be overwritten!' .
         PHP_EOL;
    sleep(10);
}

// Enable autload
require_once 'vendor/autoload.php';

// Counters
$totalRows = $totalProducts = $totalVariations = 0;

// Get list of rows
$spreadsheet = Excel\IOFactory::load($inFile);
$spreadsheet->getSheetByName('Sheet');
$rowIterator = $spreadsheet->getActiveSheet()->getRowIterator();

// Store product values to inherit in the product models (variations)
$productDataForInheritance = null;

// Iterate over rows
foreach ($rowIterator as $row) {
    $totalRows++;
    // Skip the first row with labels
    if ($row->getRowIndex() == 1) {
        continue;
    }

    // Get row cells
    $cellIterator = $row->getCellIterator();
    $isProduct    = false;

    // Iterate over the cells
    /**
     * @var Excel\Cell\Cell $cell
     */
    foreach ($cellIterator as $cellIndex => $cell) {

        // Detect if the row is a product
        if ($cellIndex === 'A') {
            if ($cell->getValue() === 'PRODUCT') {
                $isProduct = true;
                $totalProducts++;
            } elseif ($cell->getValue() === 'MODEL') {
                $totalVariations++;
            } else {
                throw new Exception(
                    'Unknown product type: ' . $cell->getValue() . ' at row ' . $row->getRowIndex()
                );
            }
        }

        // Copy columns if the row is a product. Paste columns if the column is not a product.
        if (in_array($cellIndex, ['G', 'H', 'I'])) {
            if ($isProduct) {
                $productDataForInheritance[$cellIndex] = $cell->getFormattedValue();
            } else {
                $cell->setValue($productDataForInheritance[$cellIndex]);
            }
        }

        // Optimization: do not read the rest columns
        if ($cellIndex === 'I') {
            break;
        }
    }
}

// Save modified data to a new file
$writer = new Excel\Writer\Xlsx($spreadsheet);
$writer->save($outFile);

// Print report
echo 'Done!' . PHP_EOL
     . 'Total rows: ' . $totalRows . PHP_EOL
     . 'Total products: ' . $totalProducts . PHP_EOL
     . 'Total variations: ' . $totalVariations . PHP_EOL;

exit(0);
