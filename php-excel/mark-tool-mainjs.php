<?php

/** Error reporting */
error_reporting(E_ALL);




$tool = new MarkToolMainJs();
//echo '&lt;?php <br>';
$tool->makeDdl();



class MarkToolMainJs {
    /*
 {sheet: 1, row: 12, col: 3, json: ddlSurface},
     */

    public function makeDdl() {
        for ($i=3;$i<=8;$i++){
            echo " {sheet: 1, row: 12, col: $i, json: ddlSurface},<br>";
        }
    }

   

}
