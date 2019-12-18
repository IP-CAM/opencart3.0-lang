# opencart3.0-lang
Language module for opencart 3.0. Allows to generate excel file from existing language and vise versa

This libraries can take the files from the language directory in Opencart. Then it will automatically search in all keys and export them in the specific format in excel spreadsheet. First column will contain the keys and filepath and second the value.
next you can add translation to the second column and use generation method to create a folder with corresponding files. This library guaranties that the structure of one language will be the same for the new language.

Use libraries (last version can be installed by Composer):
* phpoffice/phpspreadsheet 1.9 https://github.com/PHPOffice/PhpSpreadsheet

Has four classes and interface:
* DirectoryLister - *Search for php files and construct list of full path / Generates php files*
* ExcelException - *Exceptions of reading/writing to file*
* ExcelExporter - *Write values into excel file*
* ExcelImporter - *Read values from excel file*
* IExcelExportImport - *Contains constraint for phpspreadsheet*

Usage export:

1 Create directory list:
```php
$startDirectory = 'path to directory';
$directoryLister = new DirectoryLister('uz-uz', $startDirectory, 'php'); //'uz-uz' folder of language, php - extension of files
$files = $directoryLister->getAllFiles();
```

2 Read params from file:
```php
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

$params = readParamsFromFile($files, $startDirectory);
```

3 Populate excel file and save
```php
$excel = new ExcelExporter();
$excel->insertRow(['Kod', 'Tekst', 'Perevod']); //insert rows one-by-one
$excel->export( $startDirectory . 'demo.xlsx'); //Saves as demo.xlsx file
```

Usage import:

1. Read params from Excel:
```php
$excelimporter = new ExcelImporter($startDirectory . 'uzbek.xlsx');
$ok = $excelimporter->load();
```

2. Create folder with files:
```php
if ($ok){
    $params = $excelimporter->readRows(0,2, 2); //start column, end column and start row
    $startDirectory = 'path to directory';
    $directoryLister->populateFiles($params, $startDirectory); //Create php files in directory
}
```