<?php
require_once 'Classes/PHPExcel.php';

//create PHPExcel object
// this excel represents excel workbook
$excel = new PHPExcel();

//database connection (using mysqli)
$con = mysqli_connect("localhost", "root", "", "crs_db"); // crs_db is my database name, u must use your database name


if(!con)
{
  echo mysqli_connect_error($con);// Show what kind of errors have occured if database is not conneced
  exit;
}

//selecting active sheet
$excel->setActiveSheetIndex(0);

//populate the data
$query = mysqli_query($con,"select * from bangladesh_district_list");//bangladesh_district_list is my table name, u must use your table name

/*
  reserved three rows to
  put titles and table heading
  so I start from fourth row
*/
$row = 4;


//write your clumnNames
while ($data = mysqli_fetch_object($query)) {
  $excel->getActiveSheet()
        ->setCellValue('A'.$row, $data->id)
        ->setCellValue('B'.$row, $data->district_name)
        ->setCellValue('C'.$row, $data->status)
        ->setCellValue('D'.$row, $data->districtCode)
        ->setCellValue('E'.$row, $data->divisionCode);
        $row++;
}

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="testy.xlsx"'); // in fileName u must use your desired file name,
                                                                  //which will download as a excel file

//Every time loads a new file
header('Cache-Control: max-age=0');

//write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');

// used to cleaning html codes
ob_end_clean();

//output to php output instead of file name
$file->save('php://output');
EXIT;


/*..................IF u have a backword version of xl 2003........

......................For MsOffice xls format..change.....

header('Content-Type: application/vnd.vnd.ms-excel');
header('Content-Disposition: attachment; filename="testy.xls"');

//write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel5);

*/



?>
