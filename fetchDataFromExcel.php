<?php
require_once 'Classes/PHPExcel.php' ;

//load the excel file using PHPExcel's IOFactory where file name should be used as a parameter
$excel = PHPExcel_IOFactory::load('testy.xlsx');

//set active sheet to first sheet
$excel->setActiveSheetIndex(0);

echo "<table border=1>";


//first row of data series

//data starts from fourth row in testy file
$i = 3;

//loop until end of adata series(cell contains empty string)
while ($excel->getActiveSheet()->getCell('A'.$i)->getValue() != "") {
  # code...
  //get cell value
  $id = $excel->getActiveSheet()->getCell('A'.$i)->getValue();
  $districtName = $excel->getActiveSheet()->getCell('B'.$i)->getValue();
  $status = $excel->getActiveSheet()->getCell('C'.$i)->getValue();
  $districtCode = $excel->getActiveSheet()->getCell('D'.$i)->getValue();
  $divisionCode = $excel->getActiveSheet()->getCell('E'.$i)->getValue();

  echo " <tr>
        <td>$id.</td>
        <td>$districtName</td>
        <td>$status</td>
        <td>$districtCode</td>
        <td>$divisionCode</td>
  </tr> ";

  $i++;

}
echo "</table>";
?>
