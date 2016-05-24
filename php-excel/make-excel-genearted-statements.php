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


// RMB
$objPHPExcel->getActiveSheet()->getStyle('C19:H19')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C20:H20')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C21:H21')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C22:H22')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C23:H23')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C30:H30')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C32:H32')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C35:H35')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C36:H36')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C37:H37')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C38:H38')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C41:H41')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C44:H44')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C45:H45')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C46:H46')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C48:H48')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C51:H51')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C52:H52')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C55:H55')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C56:H56')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C59:H59')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C62:H62')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C63:H63')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C64:H64')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C67:H67')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C69:H69')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C72:H72')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C73:H73')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C76:H76')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C77:H77')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C81:H81')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C83:H83')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C87:H87')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C88:H88')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C91:H91')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C94:H94')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C95:H95')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C98:H98')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C99:H99')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C102:H102')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C103:H103')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C104:H104')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C105:H105')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C107:H107')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C108:H108')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C109:H109')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C110:H110')->getNumberFormat()->setFormatCode("¥#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C111:H111')->getNumberFormat()->setFormatCode("¥#,##0.00");

// USD
$objPHPExcel->getActiveSheet()->getStyle('C24:H24')->getNumberFormat()->setFormatCode("$#,##0.00");
$objPHPExcel->getActiveSheet()->getStyle('C112:H112')->getNumberFormat()->setFormatCode("$#,##0.00");

$objPHPExcel->getActiveSheet()
->setCellValue('C32', '=C30*C31/1000')
->setCellValue('D32', '=D30*D31/1000')
->setCellValue('E32', '=E30*E31/1000')
->setCellValue('F32', '=F30*F31/1000')
->setCellValue('G32', '=G30*G31/1000')
->setCellValue('H32', '=H30*H31/1000')
;

// --- makeColorFillStyle(A, {"A9BCF5":[15,28,39,49,53,60,65,70,74,78,84,92,96,100]})---
$objPHPExcel->getActiveSheet()->getStyle('A15')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A28')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A92')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A96')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('A100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');


// --- makeColorFillStyle(B, {"A9BCF5":[15,28,39,49,53,60,65,70,74,78,84,92,96,100]})---
$objPHPExcel->getActiveSheet()->getStyle('B15')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B28')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B39')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B49')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B53')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B60')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B78')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B84')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B92')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B96')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');
$objPHPExcel->getActiveSheet()->getStyle('B100')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('A9BCF5');


// --- makeColorFillStyle(B, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('B23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('B104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(C, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('C23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('C104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(D, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('D23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('D104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(E, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('E23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('E104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(F, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('F23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('F104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(G, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('G23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('G104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(H, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]})---
$objPHPExcel->getActiveSheet()->getStyle('H23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H38')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H48')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H52')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H59')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H64')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H69')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H73')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H77')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H83')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H91')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H95')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H99')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
$objPHPExcel->getActiveSheet()->getStyle('H104')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');
