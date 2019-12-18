<?php

require_once 'vendor/autoload.php';
require_once 'iExcelExportImport.php';
require_once 'ExcelException.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelExporter implements iExcelExportImport{
    protected $_fileName;

    private $_spreadsheet;
    private $_activesheet;
    private $_currentRow = 1;

    public function __construct($file = null)
    {
        $this->_fileName = $file;
        $this->_spreadsheet = new Spreadsheet();
        $this->_activesheet = $this->_spreadsheet->getActiveSheet();
    }

    public function getColumnAddress($column){
        if (!is_int($column)){
            throw new ExcelException('The value is not a number', ExcelException::ERROR_INDEX_NOT_A_NUMBER);
        }
        if ($column > iExcelExportImport::MAX_COLUMN_COORD){
            throw new ExcelException('The index of column is out of range', ExcelException::ERROR_INDEX_OUT_OF_RANGE);
        }
        $first = (int)log($column - 26, 26);
        $minus = 0;
        if ($first > 1){
            $minus = pow(26, $first);
        }
        $second = (int)($column - $minus) / 26;
        $third = $column % 26;
        $letterCode = '';
        $letterCode .= ($first < 2)? '' : chr($first - 2 + iExcelExportImport::SYMBOL_START_VALUE);
        $letterCode .= ($second < 1) ? (($column > 25) ? 'A': '') : chr($second - 1 + iExcelExportImport::SYMBOL_START_VALUE);
        $letterCode .= ($third < 1)? 'A' : chr($third + iExcelExportImport::SYMBOL_START_VALUE);
        return $letterCode;
    }

    public function Demo($file){
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Hello World !');

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);
    }

    public function setCell($value, $column = 'A', $row = 1){
        $column = strtoupper($column);
        $this->_activesheet->setCellValue($column . $row, $value);
    }

    public function insertRow($data, $startRow = 0, $startColumn = 0){
        $row = ($startRow == 0) ? $this->_currentRow : $startRow;
        if ( is_array($data) && (!is_array($data[0])) )
        for($iterator=0; $iterator < count($data); $iterator++ ){
            $columnLetter = $this->getColumnAddress($startColumn + $iterator);
            $this->setCell($data[$iterator], $columnLetter, $row);
        }
        $this->_currentRow++;
    }

    public function insertRows($data, $startRow = 0, $startColumn = 0){
        $row = ($startRow == 0) ? $this->_currentRow : $startRow;
        foreach ($data as $key => $value){
            //$this->_activesheet->setCellValue($column . $line, $values);
            $this->setCell($key, $this->getColumnAddress($startColumn), $row);
            $this->setCell($value, $this->getColumnAddress($startColumn+1), $row);
            $row++;
        }
        if ($row>$this->_currentRow){
            $this->_currentRow = $row - 1;
        }
    }

    public function export($fullFileName = null){
        $writer = new Xlsx($this->_spreadsheet);
        if (isset($fullFileName)) {
            $writer->save($fullFileName);
        }
        else{
            $writer->save($this->_fileName);
        }
    }
}