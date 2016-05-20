<?php

//$name = $_POST['name'];
//$city = $_POST['city'];
$data = $_POST['data']; // 傳過來的是string
$json_array = json_decode($data, true);

function getHttpToDownload($stamp) {
    $stamp = time();

//    echo "<h1>$stamp</h1>";
    $host = $_SERVER['HTTP_HOST'];
//    echo "<h1>$host</h1>";
    $script = $_SERVER['REQUEST_URI'];
//    echo "<h1>$script</h1>";

//$paths =  explode('"'+DIRECTORY_SEPARATOR+'"',$script );
//$paths =  explode(DIRECTORY_SEPARATOR,$script );
    $paths = explode('/', $script);

    $cnt = count($paths);


    $result_path = $host;
    for ($i = 0; $i < $cnt - 1; $i++) {
        $result_path.=$paths[$i] . "/";
//    echo "<h1>$result_path</h1>";
    }
    $result_path.="results/RFQ" . $stamp . ".xlsx";
//    $result_path="php-excel/results/RFQ" . $stamp . ".xlsx";
    
    
//    echo "<h1>$result_path</h1>";
    return $result_path;
//    return '<h1>'.$result_path.'</h1>';
    
}

//echo "<h1>xxx $name 在 $city </h1>";
date_default_timezone_set("Asia/Taipei");
//echo "The time is " . date("h:i:s a");
// by Mark, 2016/5/13 14:30
date_default_timezone_set("Asia/Taipei");


/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
//date_default_timezone_set('Europe/London');

define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
//require_once '../Build/PHPExcel.phar';
// NOTE by Mark 2016/5/12 3:24, 昆山
// 相信是原作者改版多次但未更新本範例
require_once '../Classes/PHPExcel.php';
//echo "__FILE__===> ".__FILE__ , EOL;
//echo "pathinfo(__FILE__, PATHINFO_BASENAME)===> ". pathinfo(__FILE__, PATHINFO_BASENAME) , EOL;
//echo "pathinfo(__FILE__, PATHINFO_DIRNAME)===> ". pathinfo(__FILE__, PATHINFO_DIRNAME) , EOL;
$stamp = time();
$defaultOutputFile = pathinfo(__FILE__, PATHINFO_DIRNAME) . DIRECTORY_SEPARATOR . "results" . DIRECTORY_SEPARATOR . "RFQ$stamp.xlsx";
//echo "\$defaultOutputFile===> ".$defaultOutputFile , EOL;
// Create new PHPExcel object
//echo date('H:i:s') , " Create new PHPExcel object" , EOL;
$objPHPExcel = new PHPExcel();

// Set document properties
//echo date('H:i:s') , " Set document properties" , EOL;
$objPHPExcel->getProperties()->setCreator("in-house WebApp RFQ")
        ->setLastModifiedBy("in-house WebApp RFQ")
        ->setTitle("RFQ")
        ->setSubject("RFQ")
        ->setDescription("RFQ generated using PHP classes.")
        ->setKeywords("office PHPExcel php")
        ->setCategory("Test result file");


// Add some data
//echo date('H:i:s') , " Add some data" , EOL;

//for ($i=0;$i<count($json_array);$i++){
    
for ($i=0;$i<32;$i++){
   $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue($json_array[$i]['pos'], $json_array[$i]['data']); 
}

$objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', $data)
        ->setCellValue('A2', "座標轉換的問題");


// Rename worksheet
//echo date('H:i:s') , " Rename worksheet" , EOL;
$objPHPExcel->getActiveSheet()->setTitle('RFQ');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Save Excel 2007 file
//echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($defaultOutputFile);
//$objWriter->save(str_replace('.php', '.xlsx', __FILE__));






$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;

//echo date('H:i:s') , " 檔案生成 " , $defaultOutputFile;
//echo $defaultOutputFile;
//echo "|";

//echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
//echo 'Check Excel file NEED TO GET PHP FILE    --- results/result123.xlsx' , EOL;
//curl_get_file_contents(getHttpToDownload($stamp));

echo getHttpToDownload($stamp);
//header("Content-disposition: attachment; filename=huge_document.pdf");
//header("Content-type: application/pdf");
//readfile("huge_document.pdf");

//
//
//echo getHttpToDownload($stamp);
////http://stackoverflow.com/questions/10088932/how-to-download-a-file-from-server-using-php-code/10088979
//function curl_get_file_contents($URL) {
//  $c = curl_init();
//  curl_setopt($c, CURLOPT_RETURNTRANSFER, 1);
//  curl_setopt($c, CURLOPT_SSL_VERIFYPEER, false);
//  curl_setopt($c, CURLOPT_URL, $URL);
//  $contents = curl_exec($c);
//  $err  = curl_getinfo($c,CURLINFO_HTTP_CODE);
//  curl_close($c);
//  if ($contents) return $contents;
//  else return FALSE;
//}