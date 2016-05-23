<?php

/** Error reporting */
error_reporting(E_ALL);




$tool = new MarkToolMainJs();
//echo '&lt;?php <br>';
$tool->makeDdl();

$tool->makeFormula23(23);
$tool->makeFormula24(24);
$tool->makeFormula30(30);
$tool->makeFormula31(31);


$tool->makeFormula29(29);
$tool->makeFormula85(85);

class MarkToolMainJs {
    /*
      {sheet: 1, row: 12, col: 3, json: ddlSurface},

      {sheet: 1, row: 10, col: 3, json: ddlMaterial},

      {sheet: 1, row: 40, col: 3, json: ddlMachine},
      {sheet: 1, row: 65, col: 3, json: ddlMaching},
      {sheet: 1, row: 70, col: 3, json: ddlCold},
      {sheet: 1, row: 74, col: 3, json: ddlSand},
      {sheet: 1, row: 79, col: 3, json: ddlStep9},
     */

    public function makeDdl() {
        for ($i = 3; $i <= 8; $i++) {
            echo " {sheet: 1, row: 12, col: $i, json: ddlSurface},<br>";
            echo "   {sheet: 1, row: 10, col: $i, json: ddlMaterial},<br>";
            echo "  {sheet: 1, row: 40, col: $i, json: ddlMachine},<br>";
            echo " {sheet: 1, row: 65, col: $i, json: ddlMaching},<br>";
            echo "  {sheet: 1, row: 70, col: $i, json: ddlCold},<br>";
            echo "  {sheet: 1, row: 74, col: $i, json: ddlSand},<br>";
            echo "  {sheet: 1, row: 79, col: $i, json: ddlStep9},<br>";
        }
    }

//  {sheet: 1, row: 29, col: 3, json: {ta: "center", dsd: "ed", data: "=C10", cal: true}},
    public function makeFormula29($row) {
//        $row=29;
        $arrAtoH = [".", "A", "B", "C", "D", "E", "F", "G", "H"];
        for ($i = 3; $i <= 8; $i++) {
            $COL = $arrAtoH[$i];
            echo "  {sheet: 1, row: $row, col: $i, json: {ta: 'center', dsd: 'ed', data: '=" . $COL . "10', cal: true}},<br>";
        }
    }

    public function makeFormula23($row) {
//        $row=23;
        $arrAtoH = [".", "A", "B", "C", "D", "E", "F", "G", "H"];
        for ($i = 3; $i <= 8; $i++) {
            $COL = $arrAtoH[$i];
//{sheet: 1, row: 23, col: 3, json: styleSubTotal({data: "=(C19+C20+C21+C22)"})},
            echo "  {sheet: 1, row: $row, col: $i, json: styleSubTotal({data: '=(" . $COL . "19+" . $COL . "20+" . $COL . "21+" . $COL . "22)'})},<br>";
        }
    }

    public function makeFormula24($row) {
//        $row=23;
        $arrAtoH = [".", "A", "B", "C", "D", "E", "F", "G", "H"];
        for ($i = 3; $i <= 8; $i++) {
            $COL = $arrAtoH[$i];
//        {sheet: 1, row: 24, col: 3, json: {fm: 'money|$|2|none', dsd: 'ed', cal: true, data: '=(C23/6.35)'}},
            $COLROW = $COL . "23";
            echo "     {sheet: 1, row: $row, col: $i, json: {fm: 'money|$|2|none', dsd: 'ed', cal: true, data: '=($COLROW/6.35)'}},<br>";
        }
    }

    public function makeFormula30($row) {
//        $row=23;
        $arrAtoH = [".", "A", "B", "C", "D", "E", "F", "G", "H"];
        for ($i = 3; $i <= 8; $i++) {
            $COL = $arrAtoH[$i];
//   {sheet: 1, row: 30, col: 3, json: {fm: 'money|¥|2|none', dsd: 'ed', cal: true, data: '=VLOOKUP(C10,LOOKUP!$A$3:$C$20,2,0)'}},
            $COLROW = $COL . "10";
            echo "     {sheet: 1, row: $row, col: $i,json: {fm: 'money|¥|2|none', dsd: 'ed', cal: true, data: '=VLOOKUP($COLROW,LOOKUP!\$A\$3:\$C\$20,2,0)'}},<br>";
        }
    }

    public function makeFormula31($row) {
        $arrAtoH = [".", "A", "B", "C", "D", "E", "F", "G", "H"];
        for ($i = 3; $i <= 8; $i++) {
            $COL = $arrAtoH[$i];
//{sheet: 1, row: 31, col: 3, json: {dsd: 'ed', data: '=C11', cal: true}},
            $COLROW = $COL . "11";
            echo "     {sheet: 1, row: $row, col: $i,json: {dsd: 'ed', data: '=$COLROW', cal: true}},<br>";
        }
    }

    public function makeFormula85($row) {
//        $row=23;
        $arrAtoH = [".", "A", "B", "C", "D", "E", "F", "G", "H"];
        for ($i = 3; $i <= 8; $i++) {
            $COL = $arrAtoH[$i];
//  {sheet: 1, row: 85, col: 3, json: {dsd: "ed", ta: "center", cal: true, data: "=C12"}},
            $COLROW = $COL . "12";
            echo "     {sheet: 1, row: $row, col: $i, json: {dsd: 'ed', ta: 'center', cal: true, data: '=$COLROW'}},<br>";
        }
    }

}
