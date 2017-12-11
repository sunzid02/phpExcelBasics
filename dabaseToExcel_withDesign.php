<?php
require_once 'Classes/PHPExcel.php';

//create PHPExcel object
// this excel represents excel workbook
$excel = new PHPExcel();

//database connection (using mysqli)
$con = mysqli_connect("localhost", "root", "", "crs_db");

if(!con)
{
  echo mysqli_connect_error($con);
  exit;
}

//selecting active sheet
$excel->setActiveSheetIndex(0);

//populate the data
$query = mysqli_query($con,"select * from bangladesh_district_list");

  /*
    reserved three rows to
    put titles and table heading
    so I start from fourth row
  */

$row = 4;

while ($data = mysqli_fetch_object($query)) {
  $excel->getActiveSheet()
        ->setCellValue('A'.$row, $data->id)
        ->setCellValue('B'.$row, $data->district_name)
        ->setCellValue('C'.$row, $data->status)
        ->setCellValue('D'.$row, $data->districtCode)
        ->setCellValue('E'.$row, $data->divisionCode);
        $row++;
}

//set column width
$excel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
$excel->getActiveSheet()->getColumnDimension('B')->setWidth(20);
$excel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
$excel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$excel->getActiveSheet()->getColumnDimension('E')->setWidth(15);

//make table headers
$excel->getActiveSheet()
        ->setCellValue('A1','Bangladesh_District_List')// this is a title
        ->setCellValue('A3','ID')
        ->setCellValue('B3','District_Name')
        ->setCellValue('C3','Status')
        ->setCellValue('D3','DistrictCode')
        ->setCellValue('E3','DivisionCode');

//merging the title
$excel->getActiveSheet()->mergeCells('A1:E1');//title range will be upto E

//aligning
$excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal('center');

//styling
$excel->getActiveSheet()->getStyle('A1')->applyFromArray(
  array(
    'font'=>array(
      'size' => 25,
    )
  )
);

$excel->getActiveSheet()->getStyle('A3:E3')->applyFromArray(
  array(
    'font'=> array(
      'bold'=> true
    )

    ,'borders'=>array(
        'allborders'=>array(
            'style'=>  PHPExcel_Style_Border::BORDER_THIN
        )
    )
  )
);

//give Borders to data
$excel->getActiveSheet()->getStyle('A4:E'.($row-1))->applyFromArray(
  array(
    'borders' => array(
      'outline' => array(
        'style' => PHPExcel_Style_Border::BORDER_THIN
      ),
      'vertical' => array(
        'style' => PHPExcel_Style_Border::BORDER_THIN
      )
    )
  )
);

/********************************Ignore**********************************************************
//insert some data to this excelsheet
/*
 before to do anything with
 the workbook,
 we have to set the
 activeWorksheet
 By calling
 setActiveSheetIndex() method

$excel->setActiveSheetIndex(0)
->setCellValue('A1','hello')
->setCellValue('B1','world');
*/
/*******************************************************************************************************/


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
ob_end_clean();
//$file->save('test.xlsx'); after loading the page file will download

//output to php output instead of file name
$file->save('php://output');
EXIT;

/*..................IF u have a bacword version of xl 2003........

......................For MsOffice xls format..change.....

header('Content-Type: application/vnd.vnd.ms-excel');
header('Content-Disposition: attachment; filename="testy.xls"');

//write the result to a file
$file = PHPExcel_IOFactory::createWriter($excel, 'Excel5);

*/



?>
