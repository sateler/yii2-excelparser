<?php

namespace sateler\excelparser;

use Exception;
use Yii;
use yii\helpers\ArrayHelper;
use yii\base\InvalidConfigException;
use yii\base\BaseObject;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\BaseReader;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;

/**
 *
 * @property-read array $data The array of model objects found
 * @property-read integer numRows The number of rows found
 * @property-read string[] missingFields The fields specified in $fields but not in the excel file
 * @property-read string[] extraFields The fields found in the excel file not in the model
 * @property-read string error The error message if one was encountered. Null otherwise
 *
 * @author felipe
 */
class ExcelParser extends BaseObject {

    /** @var string The name of the model class to create. Required. */
    public $modelClass;

    /** @var string The name of the file to load. Required if worksheet is not passed */
    public $fileName;

    /** @var string The name of the worksheet to load. Optional, used only if worksheet is not passed and fileName is used */
    public $worksheetName;

    /**
     * @var array A map of key-value pairs, where the key is the name in the
     *  excel file, and the value is the corresponding field in the model.
     *
     * Required
     */
    public $fields = [];

    /**
     * @var string[] An array of fields that are considered required: parsing will
     * not proceed if these fields are not found.
     *
     * The field name is the field in the excel sheet
     */
    public $requiredFields = [];

    /**
     * @var callable a function to determine if a row is a header row
     */
    public $isHeaderRow = null;

    /**
     * @var callable A function to create the new object
     */
    public $createObject = null;

    /** @var integer|boolean The number of rows to read at a time. Set to false (the default) to disable */
    public $chunkSize = false;

    /** @var integer|boolean Whether to save rows in an internal array to be able to retrieve it later */
    public $saveData = true;

    /**
     * @var callable A function to do something with the newly created object
     */
    public $onObjectParsed = null;

    /**
     * @var callable A function to modify the header column array before parsing data
     */
    public $modifyHeaderColumns = null;

    /** @var Worksheet */
    public $worksheet;

    /** @var boolean Wether to set values that are null */
    public $setNullValues = true;

    /** @var int The index the header row was found */
    public $headerRowIndex;

    /** @var int The index the first data row was found */
    public $dataRowIndex;

    /** @var Array */
    private $extraFields;

    /** @var Array */
    private $missingFields;

    /** @var Array */
    private $headerColumns;

    /** @var string */
    private $error = null;

    /** @var array */
    private $data = [];

    public function init() {
        if (empty($this->fields)) {
            throw new InvalidConfigException("fields is required");
        }
        if (!$this->modelClass && !$this->createObject) {
            throw new InvalidConfigException("createObject or modelClass is required");
        }
        if (!$this->fileName && !($this->worksheet instanceof Worksheet)) {
            throw new InvalidConfigException("fileName or worksheet is required");
        }

        // Sanitize fields
        $oldFields = $this->fields;
        $this->fields = [];
        foreach ($oldFields as $key => $value) {
            $this->fields[strtolower($key)] = $value;
        }
        foreach ($this->requiredFields as &$field) {
            $field = strtolower($field);
        }

        if (!$this->createObject) {
            $this->createObject = function () {
                $class = $this->modelClass;
                return new $class();
            };
        }

        if (is_null($this->isHeaderRow)) {
            $this->isHeaderRow = function () { return true; };
        }

        if (is_null($this->modifyHeaderColumns)) {
            $this->modifyHeaderColumns = function ($cols) { return $cols; };
        }

        $this->doParse();
    }

    public function getData() {
        return $this->data;
    }

    /** @return integer */
    public function getNumRows() {
        return count($this->data);
    }

    /** @return Array */
    public function getMissingFields() {
        return $this->missingFields;
    }

    /** @return Array */
    public function getExtraFields() {
        return $this->extraFields;
    }

    /** @return string */
    public function getError() {
        return $this->error;
    }

    /** @return Array */
    public function getParsedHeaders() {
        return array_keys($this->headerColumns);
    }

    private function doParse() {
        try {
            Yii::debug('Begin findHeaderRow', __CLASS__);
            $this->findHeaderRow();
            Yii::debug('Begin parseHeaderRow', __CLASS__);
            $this->parseHeaderRow();
            Yii::debug('Begin parseData', __CLASS__);
            $this->parseData();
            Yii::debug('End parseData', __CLASS__);
        }
        catch (\Exception $exc) {
            Yii::error($exc->getMessage() . "\n" . $exc->getTraceAsString(), __CLASS__);
            $this->error = $exc->getMessage();
        }
    }

    private function parseData() {
        $oldParsedData = null;
        $iter = $this->getIterator();
        $iter->startRow = $this->dataRowIndex;
        $iter->forEachRow(function ($row, $rowIndex, $sheet) use (&$oldParsedData) {
            $parsedData = call_user_func($this->createObject, $oldParsedData);
            $hasAnyValue = $this->parseRow($row, $sheet, $parsedData);
            if (!$hasAnyValue) {
                // no more data
                return false;
            }
            $oldParsedData = $parsedData;
            if($this->onObjectParsed != null ) {
                if(false === call_user_func($this->onObjectParsed, $parsedData, $rowIndex)) {
                    // stop
                    Yii::error("User returned false on onObjectParsed callable.");
                    return false;
                }
            }
            if($this->saveData) {
                $this->data[$rowIndex] = $parsedData;
            }
        });
    }

    private function parseRow(Row $row, Worksheet $sheet, &$parsedData) {
        $hasAnyValue = false;
        $rowIndex = $row->getRowIndex();
        foreach ($this->headerColumns as $key => $position) {
            $cell = $sheet->getCellByColumnAndRow($position, $rowIndex);
            $value = $cell->getCalculatedValue();
            $hasValue = !(is_null($value) || $value === '');
            $hasAnyValue = $hasAnyValue || $hasValue;
            if ($this->setNullValues || $hasValue) {
                if (Date::isDateTime($cell)) {
                    $value = Date::excelToDateTimeObject($value);
                }
                else if (is_numeric($value) && isset($parsedData->$key)) {
                    $value += $parsedData->$key;
                }
                $parsedData->$key = $value;
            }
        }
        return $hasAnyValue;
    }

    /** @return Array */
    private function parseHeaderRow() {
        $iter = $this->getIterator();
        $iter->startRow = $this->headerRowIndex;
        $iter->forEachRow(function ($headerRow) {
            $this->headerColumns = array();
            $this->extraFields = array();
            $found = array();

            $cellIterator = $headerRow->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(TRUE);
            foreach ($cellIterator as $cell) {
                $value = "" . $cell->getCalculatedValue();
                $lower = trim(strtolower($value));
                if (ArrayHelper::keyExists($lower, $this->fields, false)) {
                    // http://118.190.26.40/service/PhpSpreadsheet/commit/8d7602059006677a739c67902a434b70aa7b234c
                    $this->headerColumns[$this->fields[$lower]] = Coordinate::columnIndexFromString($cell->getColumn());
                    $found[] = $lower;
                }
                else {
                    $this->extraFields[] = $lower;
                }
            }

            $this->missingFields = array_values(array_diff(array_keys($this->fields), $found));

            $requiredMissing = array_intersect($this->missingFields, $this->requiredFields);
            if (count($requiredMissing)) {
                throw new Exception("Faltan las siguientes columnas requeridas: " . implode(", ", $requiredMissing));
            }
            return false;
        });

        $this->headerColumns = call_user_func($this->modifyHeaderColumns, $this->headerColumns);
    }

    private function findHeaderRow() {

        $iter = $this->getIterator();
        $headerCellCol = null;
        $iter->forEachRow(function ($row) use (&$headerCellCol) {
            $curCell = $this->getFirstCellWithData($row);
            if (!$curCell) {
                return;
            }
            if (!$headerCellCol && call_user_func($this->isHeaderRow, $row)) {
                $this->headerRowIndex = $curCell->getRow();
                $headerCellCol = $curCell->getColumn();
            }
            else {
                if ($headerCellCol != $curCell->getColumn()) {
                    return;
                }
                $this->dataRowIndex = $curCell->getRow();
                Yii::info("Found header row: {$this->headerRowIndex}", __CLASS__);
                return false;
            }
        });
        if (!$headerCellCol) {
            throw new Exception("El archivo tiene un formato inválido, no se pudo determinar la fila de inicio");
        }
        // have header but no data?
        $this->dataRowIndex = $this->headerRowIndex + 1;
    }

    /**
     *
     * @return \app\components\AllPseudoIterator|\app\components\ChunkedPseudoIterator
     */
    private function getIterator() {
        if ($this->worksheet) {
            return new AllPseudoIterator($this->worksheet);
        }
        else if (!$this->chunkSize) {
            Yii::debug("Begin excel open", __CLASS__);
            $reader = IOFactory::createReaderForFile($this->fileName);
            Yii::debug("Reader created", __CLASS__);
            $excel = $reader->load($this->fileName);
            Yii::debug("Excel Opened", __CLASS__);
            $this->worksheet = $this->getWorksheetFromExcelObj($excel);
            Yii::debug("End excel open", __CLASS__);
            return new AllPseudoIterator($this->worksheet);
        }
        else {
            return new ChunkedPseudoIterator($this->fileName, $this->chunkSize, $this->worksheetName);
        }
    }

    private function getWorksheetFromExcelObj(Spreadsheet $excel)
    {
        if($this->worksheetName) {
            $worksheet = $excel->getSheetByName($this->worksheetName);
            if(!$worksheet) {
                throw new Exception("No se encontró la hoja con nombre {$this->worksheetName}");
            }
            return $worksheet;
        }
        return $excel->getActiveSheet();
    }

    private function getFirstCellWithData(Row $row) {
        $cellIterator = $row->getCellIterator();
        try {
            $cellIterator->setIterateOnlyExistingCells(TRUE);
            $cellIterator->rewind();
        }
        catch (PhpSpreadsheetException $e) {
            // this happens when row is empty
            return null;
        }
        $curCell = null;
        $hasData = false;
        while (!$hasData && $cellIterator->valid()) {
            $curCell = $cellIterator->current();
            $hasData = ($curCell->getValue() !== null);
            $cellIterator->next();
        }
        return $hasData ? $curCell : null;
    }

}

class AllPseudoIterator {
    private $worksheet;

    public $startRow = 1;

    public function __construct(Worksheet $worksheet) {
        $this->worksheet = $worksheet;
    }

    public function forEachRow($function) {
        foreach ($this->worksheet->getRowIterator($this->startRow) as $row) {
            $ret = call_user_func($function, $row, $row->getRowIndex(), $this->worksheet);
            if ($ret === false) {
                break;
            }
        }
    }
}

class ChunkedPseudoIterator {

    private $chunkSize;
    private $fileName;
    /** @var ChunkReadFilter */
    private $filter;
    /** @var BaseReader */
    private $reader;

    private $sheetInfo;

    public $startRow = 1;

    public function __construct($fileName, $chunkSize, $worksheetName = null)
    {
        $this->fileName = $fileName;
        $this->chunkSize = $chunkSize;
        $this->reader = IOFactory::createReaderForFile($fileName);

        // Get sheet and row/column info
        $sheets = $this->reader->listWorksheetInfo($this->fileName);
        $this->sheetInfo = $this->getWorksheetInfo($sheets, $worksheetName);

        $this->filter = new ChunkReadFilter();
        $this->filter->setWorksheet($this->sheetInfo['worksheetName']);
        $this->reader->setReadFilter($this->filter);
    }

    private function getWorksheetInfo($sheets, $worksheetName)
    {
        if(count($sheets) <= 0) {
            throw new Exception("No se encontraron hojas con datos.");
        }
        if($worksheetName) {
            $worksheetInfos = ArrayHelper::index($sheets, 'worksheetName');
            $worksheetInfo = ArrayHelper::getValue($worksheetInfos, $worksheetName);
            if(!$worksheetInfo) {
                throw new Exception("No se encontró la hoja con nombre $worksheetName");
            }
            return $worksheetInfo;
        }
        // If no worksheet name provided just use the first one
        return $sheets[0];
    }

    public function forEachRow($function) {
        $ended = false;
        $row = $this->startRow;
        while(!$ended)
        {
            $this->filter->setRows($row, $this->chunkSize);
            //$this->filter->setMaxColumn($maxColumn);
            $this->reader->setReadFilter($this->filter);
            $spreadsheet = $this->reader->load($this->fileName);
            $spreadsheet->setActiveSheetIndexByName($this->sheetInfo['worksheetName']);
            $sheet = $spreadsheet->getActiveSheet();

            $rangeEnd = min($row + $this->chunkSize, $sheet->getHighestRow() + 1);

            for($i=$row; $i < $rangeEnd; $i++)
            {
                $objRow = $sheet->getRowIterator($i)->current();
                $ret = call_user_func($function, $objRow, $i, $sheet);
                if ($ret === false) {
                    $ended = true;
                    break;
                }
            }

            $spreadsheet->disconnectWorksheets();
            unset($spreadsheet);
            if( ($row + $this->chunkSize) > $this->sheetInfo['totalRows']) {
                $ended =true;
            }
            $row += $this->chunkSize;
        }
    }

}

class ChunkReadFilter implements IReadFilter
{
    private $_startRow = 0;
    private $_endRow = 0;
    private $_worksheetName = 0;
    private $_maxColumn = 0;
    private $_maxColumnIndex = 0;

    public function setRows($startRow, $chunkSize)
    {
        $this->_startRow    = $startRow;
        $this->_endRow      = $startRow + $chunkSize;
    }

    public function setMaxColumn($maxColumn)
    {
        $this->_maxColumn    = $maxColumn;
        $this->_maxColumnIndex = Coordinate::columnIndexFromString($maxColumn);
    }

    public function setWorksheet($worksheetName)
    {
        $this->_worksheetName = $worksheetName;
    }

    public function readCell($column, $row, $worksheetName = '')
    {
        if($this->_worksheetName == 0 || $this->_worksheetName == $worksheetName)
        {
            if( $this->_maxColumn == 0 || $this->_maxColumnIndex <=  Coordinate::columnIndexFromString($column))
            {
                if (($this->_startRow == 0) || ($row == 1) || ($row >= $this->_startRow && $row < $this->_endRow))
                {
                    return true;
                }
            }
        }
        return false;
    }
}
