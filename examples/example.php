<?php
require_once ('../vendor/autoload.php');
$dataTable = new Svrnm\ExcelDataTables\ExcelDataTable();
$in = 'spec.xlsx';

// output file need to be recreated => delete if exists
$out = 'test.xlsx';
if (file_exists($out)) {
	if (!@unlink($out)) {
		echo "CRITIC! - destination file: $out - has to be deleted, and I can't<br>";
		echo "CRITIC! - check directory and file permissions<br>";
		die();
	}
}

$data = array(
	'1' => array('A' => new DateTime()),
	'2' => array('A' => array('type' => 'datetime', 'value' => new DateTime(), 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_DATE_YYYYMMDD)),
	'3' => array('A' => array('type' => 'datetime', 'value' => new DateTime(), 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_DATE_YYYYMMDDSLASH)),
	'4' => array('A' => array('type' => 'datetime', 'value' => new DateTime(), 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_DATE_DDMMYYYY)),
	'5' => array('A' => array('type' => 'number', 'value' => 5, 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_PERCENTAGE_0)),
	'6' => array('A' => array('type' => 'number', 'value' => 10, 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_PERCENTAGE_00)),
	'7' => array('A' => array('type' => 'number', 'value' => 10.2556, 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_NUMBER_00)),
	'8' => array('A' => array('type' => 'number', 'value' => 50, 'format' => Svrnm\ExcelDataTables\ExcelNumberFormat::FORMAT_NUMBER_00)),
);
$dataTable->addRows($data)->attachToFile($in, $out, false);