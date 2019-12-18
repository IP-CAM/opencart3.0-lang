<?php
require_once 'vendor/autoload.php';
require_once 'iExcelExportImport.php';
require_once 'ExcelException.php';

use \PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelImporter implements iExcelExportImport {

    protected $_fileName;
    private $_spreadsheet;
    private $_activesheet;

    public function __construct($file = null){
        $this->_fileName = $file;
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

    public function load(){
        if (isset($this->_fileName)){
            $this->_spreadsheet = IOFactory::load($this->_fileName);
            $this->_activesheet = $this->_spreadsheet->getSheet(0);
            return true;
        }
        else {
            return false;
        }
    }

    public function getCellValue($column = 'A', $row = 1){
        return $this->_activesheet->getCell($column . $row);
    }

    public function readRows($firstColumn = 0, $secondColumn = 1, $startRow = 1){
        $params = array();
        $row = $startRow;
        while(true) {
            $key = $this->_activesheet->getCell($this->getColumnAddress($firstColumn) . $row);
            $value = $this->_activesheet->getCell($this->getColumnAddress($secondColumn) . $row);
            if ($key == '') break;
            //echo $key . ': \t' . $value;
            $params[$key . ''] = $value . '';
            $row++;
        }
        return $params;
    }
}