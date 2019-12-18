<?php
require_once '../classes/DirectoryLister.php';
require_once '../classes/ExcelExporter.php';
require_once '../classes/ExcelImporter.php';

$startDirectory = $_SERVER['DOCUMENT_ROOT'];
$directoryLister = new DirectoryLister('uz-uz', $startDirectory, 'php');

$excelimporter = new ExcelImporter($startDirectory . 'demo.xlsx');
$ok = $excelimporter->load();
if ($ok){
    $params = $excelimporter->readRows(0,2, 2);

    //Create php files in directory
    $directoryLister->populateFiles($params, $startDirectory);
}

/**
 * Read all values from php files and generate key=>value pairs
 * @param $files array array of key=>value pairs, which contains directory and file name.
 * @param $startDirectory directory directory from which start looking for php files
 * @return array Associative array key=>value pairs, which contains relative_folder*filename*paramname and paramvalue
 */
function readParamsFromFile($files, $startDirectory){
    $params = array();
    foreach ($files as $file) {
        foreach ($file as $key => $value) {
            $relativefile = $key . '/' . $value;
            $imFile = $startDirectory . '/' . $relativefile;
            include $imFile;
            $counter = 1;
            foreach ($_ as $paramName => $paramValue) {
                $params[$key . '*' . $value. '*' . $paramName] = $paramValue;
            }
            while ($value = array_pop($_)) {
            }
        }
    }
    return $params;
}