<?php

interface iExcelExportImport
{
    const SYMBOL_START_VALUE = 65;
    const SYMBOL_END_VALUE = 90;
    const MIN_COLUMN_COORD = 1;
    const MAX_COLUMN_COORD = 1000;
    const MIN_ROW_COORD = 1;
    const MAX_ROW_COORD = 1048576;
    function getColumnAddress($column);
}