<?php


$objPHPExcel->getActiveSheet()
->setCellValue('C24', '=C23/6.35')
->setCellValue('D24', '=D23/6.35')
->setCellValue('E24', '=E23/6.35')
->setCellValue('F24', '=F23/6.35')
->setCellValue('G24', '=G23/6.35')
->setCellValue('H24', '=H23/6.35')
->setCellValue('C112', '=C111/6.35')
->setCellValue('D112', '=D111/6.35')
->setCellValue('E112', '=E111/6.35')
->setCellValue('F112', '=F111/6.35')
->setCellValue('G112', '=G111/6.35')
->setCellValue('H112', '=H111/6.35')
;
$objPHPExcel->getActiveSheet()
->setCellValue('C23', '=C19+C20+C21+C22')
->setCellValue('D23', '=D19+D20+D21+D22')
->setCellValue('E23', '=E19+E20+E21+E22')
->setCellValue('F23', '=F19+F20+F21+F22')
->setCellValue('G23', '=G19+G20+G21+G22')
->setCellValue('H23', '=H19+H20+H21+H22')
->setCellValue('C111', '=C105+C107+C110')
->setCellValue('D111', '=D105+D107+D110')
->setCellValue('E111', '=E105+E107+E110')
->setCellValue('F111', '=F105+F107+F110')
->setCellValue('G111', '=G105+G107+G110')
->setCellValue('H111', '=H105+H107+H110')
->setCellValue('C110', '=C108+C109')
->setCellValue('D110', '=D108+D109')
->setCellValue('E110', '=E108+E109')
->setCellValue('F110', '=F108+F109')
->setCellValue('G110', '=G108+G109')
->setCellValue('H110', '=H108+H109')
->setCellValue('C105', '=C38+C48+C52+C59+C64+C69+C73+C77+C83+C91+C95+C99+C104')
->setCellValue('D105', '=D38+D48+D52+D59+D64+D69+D73+D77+D83+D91+D95+D99+D104')
->setCellValue('E105', '=E38+E48+E52+E59+E64+E69+E73+E77+E83+E91+E95+E99+E104')
->setCellValue('F105', '=F38+F48+F52+F59+F64+F69+F73+F77+F83+F91+F95+F99+F104')
->setCellValue('G105', '=G38+G48+G52+G59+G64+G69+G73+G77+G83+G91+G95+G99+G104')
->setCellValue('H105', '=H38+H48+H52+H59+H64+H69+H73+H77+H83+H91+H95+H99+H104')
;
$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C19:H19');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C20:H20');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C21:H21');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C22:H22');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C23:H23');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C38:H38');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C48:H48');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C52:H52');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C59:H59');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C64:H64');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C69:H69');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C73:H73');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C77:H77');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C83:H83');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C91:H91');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C95:H95');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C99:H99');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C104:H104');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C105:H105');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C110:H110');

$objPHPExcel->getActiveSheet()->duplicateConditionalStyle(
$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C111:H111');
