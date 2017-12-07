<?php
require_once 'Classes/PHPExcel.php';

//create PHPExcel object
// this excel represents excel workbook
$excel = new PHPExcel();

//insert some data to this excelsheet

/*
 before to do anything with
 the workbook,
 we have to set the
 activeWorksheet
 By calling
 setActiveSheetIndex() method
*/

$excel->setActiveSheetIndex(0)
->setCellValue('A1','hello')
->setCellValue('B1','world');

//write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$file->save('test.xlsx');

?>
