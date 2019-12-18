<?php
require_once '../classes/DirectoryLister.php';
require_once '../classes/ExcelExporter.php';
require_once '../classes/ExcelImporter.php';

$startDirectory = $_SERVER['DOCUMENT_ROOT'];
$directoryLister = new DirectoryLister('uz-uz', $startDirectory, 'php');
$files = $directoryLister->getAllFiles();

//Read params from assosiative array
$params = readParamsFromFile($files, $startDirectory);

//Write params to excel file
generateExcelFile($params, $startDirectory);

/**
 * Exports array to excel
 * @param $params array Associative array key=>value pairs, which need to write to colums of Excel file
 * @param $startDirectory directory directory where to write demo.xml file
 */
function generateExcelFile($params, $startDirectory){
    $excel = new ExcelExporter();
    $excel->insertRow(['Kod', 'Tekst', 'Perevod']);
    $excel->insertRows($params);
    $excel->export( $startDirectory . 'demo.xlsx');
}