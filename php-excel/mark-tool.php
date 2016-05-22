<?php

/** Error reporting */
error_reporting(E_ALL);




$tool = new MarkTool();
echo '&lt;?php <br>';
$tool->makeUsd();
$tool->makeSum(); //
//$tool->makeRmbStyle(); //makeRmbStyle
echo "<br> // RMB<br> ";
$moneyArrRMB = '[{"items":[19,20,21,22,23, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111]}]';
$tool->makeMoneyStyle("¥", $moneyArrRMB);
echo "<br> // USD<br> ";
$moneyArrUSD = '[{"items":[24,112]}]';
$tool->makeMoneyStyle("$", $moneyArrUSD);



class MarkTool {
    /*
      $objPHPExcel->getActiveSheet()
      ->setCellValue('C23', '=SUM(C19:C22)')
      ->setCellValue('D23', '=SUM(D19:D22)')
      ->setCellValue('E23', '=SUM(E19:E22)')
      ;
     */

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
        $strRmb = '[{"items":[19,20,21,22,23, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111]}]';
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

}
