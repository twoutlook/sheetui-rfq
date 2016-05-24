<?php

/** Error reporting */
error_reporting(E_ALL);

//$test = Array(19, 20, 21, 22, 23, 30, 32, 35, 36, 37, 38, 41, 44, 45, 46, 48, 51, 52, 55, 56, 59, 62, 63, 64, 67, 69, 72, 73, 76, 77, 81, 83, 87, 88, 91, 94, 95, 98, 99, 102, 103, 104, 105, 107, 108, 109, 110, 111);


$tool = new MarkTool();
echo '&lt;?php <br>';
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


//    var colorStep = "#A9BCF5";
//    var colorStepEnd = "#E6E6E6";
//    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
//    var colorDdl = "#F9E79F"; //#82E0AA  
//    var colorInput = "#F4D03F"; // 
//    var arrStepEnd = [23, 24, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111, 112];
$colorJsonStrStepStart = '{"A9BCF5":[15,28,39,49,53,60,65,70,74,78,84,92,96,100]}';
$colorJsonStrStepEnd = '{"BE6E6E6":[23,38,48,52,59,64,69,73,77,83,91,95,99,104]}';

$tool->makeColorFillStyle("A", $colorJsonStrStepStart);
$tool->makeColorFillStyle("B", $colorJsonStrStepStart);
$tool->makeColorFillStyle("B", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("C", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("D", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("E", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("F", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("G", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("H", $colorJsonStrStepEnd);


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
