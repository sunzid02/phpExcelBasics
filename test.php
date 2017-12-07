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

// redirect to browser (download) instead of saving the result as a get_included_file
/*
  First we need to setup the http header
  to tell the browser that this is an excel
  speradsheet file, not a regular html file
*/
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="testy.xlsx"');

//Every time loads a new file
header('Cache-Control: max-age=0');

//write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');

//$file->save('test.xlsx'); after loading the page file will download

//output to php output instead of file name
$file->save('php://output');


/*..................IF u have a bacword version of xl 2003........

......................For MsOffice xls format..change.....

header('Content-Type: application/vnd.vnd.ms-excel');
header('Content-Disposition: attachment; filename="testy.xls"');

//write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel5);

*/



?>
