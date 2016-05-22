<?php

/** Error reporting */
error_reporting(E_ALL);




$tool = new MarkTool();
$tool->makeUsd();
$tool->makeSum();


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
        print_r($objUsd);

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
        $strUsd = '[{"sum":23, "items":[19,20,21,22]}]';
        $objUsd = json_decode($strUsd);
        print_r($objUsd);

        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
        foreach ($objUsd as $key => $obj) {
            $colNameArr = Array("C", "D", "E", "F", "G", "H");
            for ($i = 0; $i < 6; $i++) {
                $colName = $colNameArr[$i];
//                $str.="  ->setCellValue('" . $colName . $obj->sum . "', '=" . $colName . $obj->rmb . "/6.35') <br>";
                $str.="  ->setCellValue('" . $colName . $obj->sum . "', '=" ;
                
                $items=$obj->items;
                for ($j = 0; $j < count($items) ; $j++){
                    $str.= $colName .$items[$j];
                    if ($j<count($items)-1){
                        $str.="+"; 
                    }
                }
                $str.="')<br>";
            }
        }
        echo $str . ";";
    }
}
