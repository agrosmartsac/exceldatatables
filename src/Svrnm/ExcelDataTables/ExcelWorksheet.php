<?php

namespace Svrnm\ExcelDataTables;

/**
 * An instance of this class represents a simple(!) ExcelWorkseeht in the spreadsheetml format.
 * The most important function ist the addRow() function which takes an array as parameter and
 * adds its values to the worksheet. Finally the worksheet can be exported to XML using the toXML()
 * method
 *
 * @author Severin Neumann <s.neumann@altmuehlnet.de>
 * @license Apache-2.0 
 */
class ExcelWorksheet
{
	/**
	 * This namespaces are used to setup the XML document.
	 *
	 * @var array
	 */
	protected static $namespaces = array(
		"spreadsheets" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		"relationships" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		"xmlns" => "http://www.w3.org/2000/xmlns/"
	);

	/**
	 * The base date which is used to compute date field values
	 *
	 * @var string
	 */
	protected static $baseDate = "1899-12-31 00:00:00";

	protected static $fixDays = 1;

	/**
	 * The XML base document
	 *
	 * @var \DOMDocument
	 */
	protected $document;

	/**
	 * The worksheet element. This is the root element of the XML document
	 *
	 * @var \DOMElement
	 */
	protected $worksheet;
	/**
	 * The sheetData element. This element contains all rows of the spreadsheet
	 *
	 * @var \DOMElement
	 */
	protected $sheetData;

	protected $sharedStrings;

	/**
	 * The formatId used for date and time values. The correct id is specified
	 * in the styles.xml of a workbook. The default value 1 is a placeholder
	 *
	 * @var int
	 */
	protected $dateTimeFormatId = 1;

	protected $dateTimeColumns = array();

	protected $rows = array();

	protected $dirty = false;

	protected $rowCounter = 1;

	protected $usedFormats = array();

	const COLUMN_TYPE_STRING = 0;
	const COLUMN_TYPE_NUMBER = 1;
	const COLUMN_TYPE_DATETIME = 2;
	const COLUMN_TYPE_FORMULA = 3;
	const COLUMN_TYPE_SHAREDSTRING = 4;

	protected static $columnTypes = array(
		'string' => 0,
		'number' => 1,
		'datetime' => 2,
		'formula' => 3,
		'sharedstring' => 4
	);

	protected $oldXml = '';
	protected $preserveRows = array();

	/**
	 * Setup a default document: XML head, Worksheet element, SheetData element.
	 *
	 * @return $this
	 */
	public function setupDefaultDocument()
	{
		$this->getSheetData();
		return $this;
	}

	/**
	 * Change the formatId for date time values.
	 *
	 * @return $this
	 */
	public function setDateTimeFormatId($id)
	{
		$this->dirty = true;
		$this->dateTimeFormatId = $id;
		return $this;
	}

	/**
	 * Convert DateTime to excel time format. This function is
	 * a copy from PHPExcel.
	 *
	 * @see https://github.com/PHPOffice/PHPExcel/blob/78a065754dd0b233d67f26f1ef8a8a66cd449e7f/Classes/PHPExcel/Shared/Date.php
	 */
	public static function convertDate(\DateTime $date)
	{

		$year = $date->format('Y');
		$month = $date->format('m');
		$day = $date->format('d');
		$hours = $date->format('H');
		$minutes = $date->format('i');
		$seconds = $date->format('s');

		$excel1900isLeapYear = TRUE;
		if (($year == 1900) && ($month <= 2)) {
			$excel1900isLeapYear = FALSE;
		}
		$my_excelBaseDate = 2415020;
		if ($month > 2) {
			$month -= 3;
		} else {
			$month += 9;
			$year -= 1;
		}
		// Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
		$century = substr($year, 0, 2);
		$decade = substr($year, 2, 2);
		$excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5) + $day + 1721119 - $my_excelBaseDate + $excel1900isLeapYear;

		$excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

		return (float) $excelDate + $excelTime;

	}

	/**
	 * By default the XML document is generated without format. This can be
	 * changed with this function.
	 *
	 * @param $value
	 * @return $this
	 */
	public function setFormatOutput($value = true)
	{
		$this->getDocument()->formatOutput = true;
		return $this;
	}

	/**
	 * Returns the given worksheet in its XML representation
	 *
	 * @return string
	 */
	public function toXML()
	{
		$document = $this->getDocument();
		return $document->saveXML();
	}

	/**
	 * Generate and return a new empty row within the sheetData
	 *
	 * @return \DOMElement
	 */
	protected function getNewRow()
	{
		$this->rows[] = array();
		return count($this->rows) - 1;
	}

	/**
	 * Add a column to a row. The type of the column is deferred by its value
	 *
	 * @param \DOMElement $row
	 * @param mixed $column
	 * @return \DOMElement
	 */
	protected function addColumnToRow($row, $column, $data)
	{
		if (
			is_array($data)
			&& isset($data['type'])
			&& isset($data['value'])
			&& in_array($data['type'], array('string', 'number', 'datetime', 'formula', 'sharedstring'))
		) {
			if (in_array($data['type'], array('number', 'datetime')) && !empty($data['format'])) {
				$this->usedFormats[$data['format']] = -1;
			} else {
				$data['format'] = "";
			}
			$this->rows[$row][$column] = array(self::$columnTypes[$data['type']], $data['value'], ($data['format'] ?? ""));
		} elseif (is_numeric($data)) {
			$this->rows[$row][$column] = array(self::COLUMN_TYPE_NUMBER, $data, "");
		} elseif ($data instanceof \DateTime) {
			$this->usedFormats[ExcelNumberFormat::FORMAT_DATE_DEFAULT] = -1;
			$this->rows[$row][$column] = array(self::COLUMN_TYPE_DATETIME, $data, ExcelNumberFormat::FORMAT_DATE_DEFAULT);
		} else {
			$this->rows[$row][$column] = array(self::COLUMN_TYPE_STRING, $data, "");
		}
	}

	public function toXMLColumn($row, $column, $data)
	{
		$formatId = (!empty($data[2]) ? ($this->usedFormats[$data[2]] ?? -1) : -1);
		switch ($data[0]) {
			case self::COLUMN_TYPE_NUMBER:
				if ($formatId == -1) {
					return '<c r="' . $column . $row . '"><v>' . $data[1] . '</v></c>';
				} else {
					return '<c r="' . $column . $row . '" s="' . $formatId . '"><v>' . $data[1] . '</v></c>';
				}
				break;
			case self::COLUMN_TYPE_DATETIME:
				return '<c r="' . $column . $row . '" s="' . $formatId . '"><v>' . static::convertDate($data[1]) . '</v></c>';
				break;
			case self::COLUMN_TYPE_FORMULA:
				return '<c r="' . $column . $row . '"><f>' . $data[1] . '</f></c>';
				break;
			case self::COLUMN_TYPE_SHAREDSTRING:
				return '<c r="' . $column . $row . '" t="s"><v>' . $data[1] . '</v></c>';
				break;
			default:
				return '<c r="' . $column . $row . '" t="inlineStr"><is><t>' . strtr(
					$data[1],
					array(
						"&" => "&amp;",
						"<" => "&lt;",
						">" => "&gt;",
						'"' => "&quot;",
						"'" => "&apos;",
					)
				) . '</t></is></c>';
				break;
		}
	}

	public function incrementRowCounter()
	{
		return $this->rowCounter++;
	}

	protected function updateDocument()
	{
		if ($this->dirty) {
			$this->dirty = false;
			$self = $this;
			$this->rowCounter = 1;
			$fragment = $this->document->createDocumentFragment();
			$xml = $this->oldXml;
			foreach ($this->rows as $row => $rowData) {
				if (empty($rowData) && isset($this->preserveRows[$row])) {
					$xml .= $this->preserveRows[$row];
				} else {
					$xml .= '<row r="' . $row . '">';
					foreach ($rowData as $column => $columnData) {
						$xml .= $this->toXMLColumn($row, $column, $columnData);
					}
					$xml .= '</row>';
				}
			}
			if (!$fragment->appendXML($xml)) {
				throw new \Exception('Parsing XML failed.');
			}
			$this->getSheetData()->parentNode->replaceChild(
				$s = $this->getSheetData()->cloneNode(false),
				$this->getSheetData()
			);
			$this->sheetData = $s;
			$this->getSheetData()->appendChild($fragment);

		}
	}

	/**
	 * Add a row to the spreadsheet. The columns are inserted and their type is deferred by their type:
	 *
	 * - Arrays having a type and value element are inserted as defined by the type. Possible types
	 * are: string, number, datetime
	 * - Numerical values are inserted as number columns.
	 * - Objects implementing the DateTimeInterface are inserted as datetime column.
	 * - Everything else is converted to a string and inserted as (inline) string column.
	 *
	 * @param array $columns
	 * @return $this
	 */
	public function addRow($row, $columns = array())
	{
		$this->dirty = true;
		//$row = $this->getNewRow();
		if (!isset($this->rows[$row])) {
			$this->rows[$row] = array();
		}
		foreach ($columns as $column => $data) {
			$this->addColumnToRow($row, $column, $data);
		}
		return $this;
	}

	/**
	 * Returns the DOMDocument representation of the current instance
	 *
	 * @return \DOMDocument
	 */
	public function getDocument()
	{
		if (is_null($this->document)) {
			$this->document = new \DOMDocument('1.0', 'utf-8');
			$this->document->xmlStandalone = true;
		}
		$this->updateDocument();
		return $this->document;
	}

	/**
	 * Returns the DOMElement representation of the sheet data
	 *
	 * @return \DOMElement
	 */
	public function getSheetData()
	{
		if (is_null($this->sheetData)) {
			$this->sheetData = $this->append('sheetData');
		}
		$this->updateDocument();
		return $this->sheetData;
	}

	/**
	 * Crate a new \DOMElement within the scope of the current document.
	 *
	 * @param string name
	 * @return \DOMElement
	 */
	protected function createElement($name)
	{
		return $this->getDocument()->createElementNS(static::$namespaces['spreadsheets'], $name);
	}

	/**
	 * Returns the DOMElement representation of the worksheet
	 *
	 * @return \DOMElement
	 */
	public function getWorksheet()
	{
		if (is_null($this->worksheet)) {
			$document = $this->getDocument();
			$this->worksheet = $this->append('worksheet', array(), $document);
			$this->worksheet->setAttributeNS(static::$namespaces['xmlns'], 'xmlns:r', static::$namespaces['relationships']);
		}
		$this->updateDocument();
		return $this->worksheet;
	}

	/**
	 * Append a new element (tag) to the XML Document. By default the new tag <$name/> will be attachted
	 * to the root element (i.e. <worksheet>). Attributes for the new tag can be specified with the second
	 * parameter $attribute. Each element of the $attributes array is added as attribute whereas the key
	 * is the attribute name and the value is the attribute value.
	 * If the new element should be appended to another parent element in the XML Document the third
	 * parameter can be used to specify the parent
	 *
	 * The function returns the newly created element as \DOMElement instance.
	 *
	 * @param string name
	 * @param array attributes
	 * @param \DOMElement parent
	 * @return \DOMElement
	 */
	protected function append($name, $attributes = array(), $parent = null)
	{
		if (is_null($parent)) {
			$parent = $this->getWorksheet();
		}
		$element = $this->createElement($name);
		foreach ($attributes as $key => $value) {
			$element->setAttribute($key, $value);
		}
		$parent->appendChild($element);
		return $element;
	}

	public function addRows($array, $calculatedColumns = null)
	{
		foreach ($array as $key => $row) {
			if (isset($calculatedColumns)) {
				foreach ($calculatedColumns as $calculatedColumn) {
					if ($key == 0) {
						array_splice($row, $calculatedColumn['index'], 0, $calculatedColumn['header']);
					} else {
						array_splice($row, $calculatedColumn['index'], 0, $calculatedColumn['content']);
					}
				}
			}
			$this->addRow($key, $row);
		}
		return $this;
	}

	public function clearPreserveRows()
	{
		$this->preserveRows = array();
	}

	public function setPreserveRow($rowNumber, $xmlElement)
	{
		$this->preserveRows[$rowNumber] = $xmlElement;
	}

	public function setOldXml($oldXml)
	{
		$this->oldXml = $oldXml;
	}

	public function setUsedFormats($usedFormats)
	{
		$this->usedFormats = $usedFormats;
	}

	public function getUsedFormats()
	{
		return $this->usedFormats;
	}

}
