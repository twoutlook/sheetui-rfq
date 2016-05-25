<?php

/** Error reporting */
error_reporting(E_ALL);

//$test = Array(19, 20, 21, 22, 23, 30, 32, 35, 36, 37, 38, 41, 44, 45, 46, 48, 51, 52, 55, 56, 59, 62, 63, 64, 67, 69, 72, 73, 76, 77, 81, 83, 87, 88, 91, 94, 95, 98, 99, 102, 103, 104, 105, 107, 108, 109, 110, 111);


$tool = new MarkTool();
echo '&lt;?php <br>';
$tool->makePercentFormat();
$tool->makeUsd();
$tool->makeSum(); //
//$tool->makeRmbStyle(); //makeRmbStyle
echo "<br> // RMB<br> ";
$moneyArrRMB = '[{"items":[19, 20, 21, 22, 23, 30, 32, 35, 36, 37, 38, 41, 44, 45, 46, 48, 51, 52, 55, 56, 59, 62, 63, 64, 67, 69, 72, 73, 76, 77, 81, 83, 87, 88, 91, 94, 95, 98, 99, 102, 103, 104, 105, 107, 108, 109, 110, 111]}]';
$tool->makeMoneyStyle("¥", $moneyArrRMB);
echo "<br> // USD<br> ";
$moneyArrUSD = '[{"items":[24,112]}]';
$tool->makeMoneyStyle("$", $moneyArrUSD);
$tool->makeCell32(32);
$tool->extendCell34X(34, "=100*C31*C16/(C31*C16+C33)/100"); //注意 EXCEL 和 ENTERPRISESHEET 的百分比表達方式不同
$tool->extendCell34X(36, "=(C30-C35)*C33/1000/C16");
$tool->extendCell34X(37, "=(C31+C33)*C30*0.02/1000/C16");
$tool->extendCell34X(38, "=IF(ISNA(C32+C36+C37),0,(C32+C36+C37))");
//
$tool->extendCell34X(43, "=3600/C42");
$tool->extendCell34X(44, "=C41/C43 ");
//$tool->extendCell34X(45, "  ");
//$tool->extendCell34X(46, "  ");
// 47 百分比要處理
// $tool->extendCell34X(48, "=(C44+C45)*(1+(1-C47/100))/C16");  // before fix
$tool->extendCell34X(48, "=(C44+C45)*(1+(1-C47))/C16");  //  after fix

$tool->extendCell34X(52, "=(C50/3600)*C51");
//
$tool->extendCell34X(56, "=(C54/3600)*C55");
//=C56*(1+(1-C57/100))
//$tool->extendCell34X(59, "=C56*(1+(1-C57/100))"); // TO FIX PERCENT
$tool->extendCell34X(59, "=C56*(1+(1-C57))");

// 63,64 SAME =(C61/3600)*C62'
$tool->extendCell34X(63, "=(C61/3600)*C62");
$tool->extendCell34X(64, "=(C61/3600)*C62");
//
//$tool->extendCell34X(69, "=(C66/3600)*C67 * (1 + (1 - C68 / 100))");
$tool->extendCell34X(69, "=IF(ISNA((C66/3600)*C67 * (1 + (1 - C68))),0,(C66/3600)*C67 * (1 + (1 - C68)))");

//73
//(C71/3600)*C72
$tool->extendCell34X(73, "=IF(ISNA((C71/3600)*C72),0,(C71/3600)*C72)");

//77
//"=IF(ISNA(C76),0,(C75/3600)*C76)
$tool->extendCell34X(77, "=IF(ISNA(C76),0,(C75/3600)*C76)");

//83
//
$tool->extendCell34X(83, "=C80*C81*C82");
//91
//"=C86*(C87+C88)*(1+(1-C89/100))*C90
//$tool->extendCell34X(91, "=C86*(C87+C88)*(1+(1-C89/100))*C90");
$tool->extendCell34X(91, "=C86*(C87+C88)*(1+(1-C89))*C90");

//95
//
$tool->extendCell34X(95, "=C94");

//99
//
$tool->extendCell34X(99, "=C98");

//104
//
$tool->extendCell34X(104, "=C102+C103");

//107
//
//$tool->extendCell34X(107, "=C105*C106/100");
$tool->extendCell34X(107, "=C105*C106");


//110
//
$tool->extendCell34X(110, "=C108+C109");




//
//
//
//
//
//    var colorStep = "#A9BCF5";
//    var colorStepEnd = "#E6E6E6";
//    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
//    var colorDdl = "#F9E79F"; //#82E0AA  
//    var colorInput = "#F4D03F"; // 
//    var arrStepEnd = [23, 24, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111, 112];
$colorJsonStrStepStart = '{"A9BCF5":[15,28,39,49,53,60,65,70,74,78,84,92,96,100]}';
$colorJsonStrStepEnd = '{"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104,110]}';
$jsonDdl = '{"F9E79F":[10,12,29,40,65,70,74,79,85]}';
$jsonInput = '{"F4D03F":[11,16,19,20,21,22,33,42,47,50,54,57,61,66,68,71,75,80,81,82,86,87,88,89,90,93,94,97,101,103,106,108,109]}';
$jsonBigTotal = '{"cccccc":[23,24,105,111,112]}';

$tool->makeColorFillStyle("A", $colorJsonStrStepStart);
$tool->makeColorFillStyle("B", $colorJsonStrStepStart);
$tool->makeColorFillStyle("B", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("C", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("D", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("E", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("F", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("G", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("H", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("C", $jsonDdl);
$tool->makeColorFillStyle("D", $jsonDdl);
$tool->makeColorFillStyle("E", $jsonDdl);
$tool->makeColorFillStyle("F", $jsonDdl);
$tool->makeColorFillStyle("G", $jsonDdl);
$tool->makeColorFillStyle("H", $jsonDdl);
$tool->makeColorFillStyle("C", $jsonInput);
$tool->makeColorFillStyle("D", $jsonInput);
$tool->makeColorFillStyle("E", $jsonInput);
$tool->makeColorFillStyle("F", $jsonInput);
$tool->makeColorFillStyle("G", $jsonInput);
$tool->makeColorFillStyle("H", $jsonInput);

$tool->makeColorFillStyle("A", $jsonBigTotal);
$tool->makeColorFillStyle("B", $jsonBigTotal);
$tool->makeColorFillStyle("C", $jsonBigTotal);
$tool->makeColorFillStyle("D", $jsonBigTotal);
$tool->makeColorFillStyle("E", $jsonBigTotal);
$tool->makeColorFillStyle("F", $jsonBigTotal);
$tool->makeColorFillStyle("G", $jsonBigTotal);
$tool->makeColorFillStyle("H", $jsonBigTotal);

class MarkTool {
    /*
      $objPHPExcel->getActiveSheet()
      ->setCellValue('C23', '=SUM(C19:C22)')
      ->setCellValue('D23', '=SUM(D19:D22)')
      ->setCellValue('E23', '=SUM(E19:E22)')
      ;
     */

    public function makeCell32($row) {
        $colNameArr = Array("C", "D", "E", "F", "G", "H");
        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
        for ($i = 0; $i < 6; $i++) {
            $colName = $colNameArr[$i];
            //=C30*C31/1000
            $colRow = $colName . $row;
            $cell = "=" . $colName . "30*" . $colName . "31/1000";
            $str.="  ->setCellValue('$colRow', '$cell') <br>";
        }
        echo $str . ";";
    }

    //=100*C31*C16/(C31*C16+C33)
//      public function makeCell34X($rowNum,$cellFormula) {
//        $colNameArr = Array("C", "D", "E", "F", "G", "H");
//        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
//        for ($i = 0; $i < 6; $i++) {
//            $colName = $colNameArr[$i];
//            //=C30*C31/1000
//            $colRow = $colName . $row;
//            $cell = "=" . $colName . "30*" . $colName . "31/1000";
//            $str.="  ->setCellValue('$colRow', '$cell') <br>";
//        }
//        echo $str . ";";
//    }


    public function extendCell34X($rowNum, $cellFormula) {
        echo "<br><br>// ---  extendCell34X($rowNum,$cellFormula) ---<br>";

        $seq = "0ABCDEFGH";
        echo " <br> \$objPHPExcel->getActiveSheet() <br>";
        echo "->setCellValue('C$rowNum', '$cellFormula')<br>";
        for ($k = 4; $k <= 8; $k++) {
            //    $strD = str_replace("col: 3", "col: $k", $cellFormula);
            $colName = substr($seq, $k, 1);
            $newCellformula = $cellFormula;
            for ($item = 1; $item < 115; $item++) {
                $oldStr = "C$item";
                $newStr = $colName . $item;
                $newCellformula = str_replace($oldStr, $newStr, $newCellformula);
            }
            echo "->setCellValue('$colName$rowNum', '$newCellformula')<br>";
        }
        echo ";<br>";
    }

    public function makeUsd() {
        $strUsd = '[{"usd":24, "rmb":23},{"usd":112, "rmb":111}]';
        $objUsd = json_decode($strUsd);
//        print_r($objUsd);

        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
        foreach ($objUsd as $key => $obj) {
            $colNameArr = Array("C", "D", "E", "F", "G", "H");
            for ($i = 0; $i < 6; $i++) {
                $colName = $colNameArr[$i];
                $str.="  ->setCellValue('" . $colName . $obj->usd . "', '=" . $colName . $obj->rmb . "/6.35') <br>";
            }
        }
        echo $str . ";";
    }

    public function makeSum() {
        $strUsd = '[{"sum":23, "items":[19,20,21,22]},{"sum":111, "items":[105,107,110]},{"sum":110, "items":[108,109]},{"sum":105, "items":[38,48,52,59,64,69,73,77,83,91,95,99,104]}]';
        $objUsd = json_decode($strUsd);
//        print_r($objUsd);

        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
        foreach ($objUsd as $key => $obj) {
            $colNameArr = Array("C", "D", "E", "F", "G", "H");
            for ($i = 0; $i < 6; $i++) {
                $colName = $colNameArr[$i];
//                $str.="  ->setCellValue('" . $colName . $obj->sum . "', '=" . $colName . $obj->rmb . "/6.35') <br>";
                $str.="  ->setCellValue('" . $colName . $obj->sum . "', '=";

                $items = $obj->items;
                for ($j = 0; $j < count($items); $j++) {
                    $str.= $colName . $items[$j];
                    if ($j < count($items) - 1) {
                        $str.="+";
                    }
                }
                $str.="')<br>";
            }
        }
        echo $str . ";<br><br>";
    }

    /*

      $objPHPExcel->getActiveSheet()->getStyle('C19:H23')->getNumberFormat()->setFormatCode("¥#,##0.00");
      $objPHPExcel->getActiveSheet()->getStyle('C24:H24')->getNumberFormat()->setFormatCode("$#,##0.00");
     * * 
     */

    public function makeRmbStyle() {
        $strRmb = '[{"items":[19,20,21,22,23,32, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111]}]';
        $objRmb = json_decode($strRmb);
//        print_r($objRmb);



        foreach ($objRmb as $key => $obj) {

            $items = $obj->items;
            for ($j = 0; $j < count($items); $j++) {
//                $str = " <br> \$objPHPExcel->getActiveSheet()->duplicateConditionalStyle( <br>";
//                $str.= "  \$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C" . $items[$j] . ":H" . $items[$j];
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('C" . $items[$j] . ":H" . $items[$j] . "')->getNumberFormat()->setFormatCode(\"¥#,##0.00\")";
                echo $str . ";<br>";
            }
        }
    }

    public function makePercentFormat() {
        echo "<br>// makePercentFormat 0.00% ok, FORMAT_PERCENTAGE_00 為什麼不行?<br>";
        $strRmb = '[{"items":[34,47,57,68,89,106]}]';
        $objRmb = json_decode($strRmb);

        foreach ($objRmb as $key => $obj) {

            $items = $obj->items;
            for ($j = 0; $j < count($items); $j++) {
//                $str = " <br> \$objPHPExcel->getActiveSheet()->duplicateConditionalStyle( <br>";
//                $str.= "  \$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C" . $items[$j] . ":H" . $items[$j];
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('C" . $items[$j] . ":H" . $items[$j] . "')->getNumberFormat()->setFormatCode(\"0.00%\")";
                echo $str . ";<br>";
            }
        }
    }

    public function makeMoneyStyle($money, $moneyArr) {
        $objMoney = json_decode($moneyArr);
//        print_r($objRmb);



        foreach ($objMoney as $key => $obj) {

            $items = $obj->items;
            for ($j = 0; $j < count($items); $j++) {
//                $str = " <br> \$objPHPExcel->getActiveSheet()->duplicateConditionalStyle( <br>";
//                $str.= "  \$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C" . $items[$j] . ":H" . $items[$j];
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('C" . $items[$j] . ":H" . $items[$j] . "')->getNumberFormat()->setFormatCode(\"" . $money . "#,##0.00\")";
                echo $str . ";<br>";
            }
        }
    }

    // cell format
    // http://www.cnblogs.com/freespider/p/3284828.html
    // for column B only
    public function makeColorFillStyle($col, $str) {
        echo "<br><br>// ---  makeColorFillStyle($col, $str)---<br> ";
        $json = json_decode($str);
//        print_r($json);
        //$objPHPExcel->getActiveSheet()->getStyle('A1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
//        ->getStartColor()->setARGB('FFA9BCF5');

        foreach ($json as $key => $val) {
//            echo $key;
//            print_r($val);
            foreach ($val as $item) {
//                 echo $item;
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('$col$item')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)";
                $str.=" ->getStartColor()->setARGB('$key');<br>";
                echo $str;
            }
        }
    }

}
