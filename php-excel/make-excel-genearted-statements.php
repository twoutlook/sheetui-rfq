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


// --- makeColorFillStyle(B, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('B110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(C, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('C110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(D, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('D110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(E, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('E110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(F, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('F110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(G, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('G110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(H, {"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]})---
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
$objPHPExcel->getActiveSheet()->getStyle('H110')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('BE6E6E6');


// --- makeColorFillStyle(C, {"F9E79F":[10,12,29,40,65,70,74,79,85]})---
$objPHPExcel->getActiveSheet()->getStyle('C10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C40')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C79')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('C85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


// --- makeColorFillStyle(D, {"F9E79F":[10,12,29,40,65,70,74,79,85]})---
$objPHPExcel->getActiveSheet()->getStyle('D10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D40')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D79')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('D85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


// --- makeColorFillStyle(E, {"F9E79F":[10,12,29,40,65,70,74,79,85]})---
$objPHPExcel->getActiveSheet()->getStyle('E10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E40')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E79')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('E85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


// --- makeColorFillStyle(F, {"F9E79F":[10,12,29,40,65,70,74,79,85]})---
$objPHPExcel->getActiveSheet()->getStyle('F10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F40')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F79')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('F85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


// --- makeColorFillStyle(G, {"F9E79F":[10,12,29,40,65,70,74,79,85]})---
$objPHPExcel->getActiveSheet()->getStyle('G10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G40')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G79')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('G85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


// --- makeColorFillStyle(H, {"F9E79F":[10,12,29,40,65,70,74,79,85]})---
$objPHPExcel->getActiveSheet()->getStyle('H10')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H12')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H29')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H40')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H65')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H70')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H74')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H79')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');
$objPHPExcel->getActiveSheet()->getStyle('H85')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F9E79F');


// --- makeColorFillStyle(C, {"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]})---
$objPHPExcel->getActiveSheet()->getStyle('C11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C20')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C21')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C22')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C42')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C68')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C81')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C82')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C87')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C88')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C89')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C90')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C93')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C94')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C97')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C101')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C106')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('C109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


// --- makeColorFillStyle(D, {"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]})---
$objPHPExcel->getActiveSheet()->getStyle('D11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D20')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D21')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D22')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D42')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D68')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D81')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D82')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D87')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D88')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D89')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D90')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D93')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D94')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D97')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D101')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D106')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('D109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


// --- makeColorFillStyle(E, {"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]})---
$objPHPExcel->getActiveSheet()->getStyle('E11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E20')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E21')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E22')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E42')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E68')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E81')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E82')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E87')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E88')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E89')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E90')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E93')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E94')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E97')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E101')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E106')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('E109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


// --- makeColorFillStyle(F, {"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]})---
$objPHPExcel->getActiveSheet()->getStyle('F11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F20')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F21')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F22')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F42')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F68')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F81')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F82')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F87')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F88')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F89')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F90')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F93')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F94')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F97')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F101')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F106')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('F109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


// --- makeColorFillStyle(G, {"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]})---
$objPHPExcel->getActiveSheet()->getStyle('G11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G20')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G21')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G22')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G42')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G68')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G81')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G82')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G87')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G88')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G89')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G90')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G93')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G94')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G97')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G101')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G106')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('G109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


// --- makeColorFillStyle(H, {"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]})---
$objPHPExcel->getActiveSheet()->getStyle('H11')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H16')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H19')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H20')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H21')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H22')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H33')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H42')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H47')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H50')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H54')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H57')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H61')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H66')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H68')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H71')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H75')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H80')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H81')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H82')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H86')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H87')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H88')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H89')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H90')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H93')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H94')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H97')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H101')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H103')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H106')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H108')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');
$objPHPExcel->getActiveSheet()->getStyle('H109')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('F4D03F');


// --- makeColorFillStyle(A, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('A23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('A112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(B, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('B23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('B112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(C, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('C23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('C112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(D, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('D23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('D112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(E, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('E23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('E112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(F, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('F23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('F112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(G, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('G23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('G112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');


// --- makeColorFillStyle(H, {"cccccc":[23,24,105,111,112]})---
$objPHPExcel->getActiveSheet()->getStyle('H23')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H24')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H105')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H111')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
$objPHPExcel->getActiveSheet()->getStyle('H112')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID) ->getStartColor()->setARGB('cccccc');
