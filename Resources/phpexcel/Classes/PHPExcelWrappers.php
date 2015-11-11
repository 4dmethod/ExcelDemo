<?php
require_once 'PHPExcel.php';
require_once 'PHPExcel/IOFactory.php';

/*******************************
* GENERAL NOTES:
* --------------
* Created by Brent Raymond
* 4DMethod User Group, 4dmethod.com
* see Tech Note: PHPExcel Library with 4D v12, http://kb.4d.com/assetid=76312
* also see http://stackoverflow.com/questions/3537604/how-to-fix-a-memory-error-in-php
********************************/

function excel_Upvert($excelFile){
	// initialize cell caching settings to conserve memory and be able to handle larger files
	$cacheMethod = PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
	$cacheSettings = array( ' memoryCacheSize ' => '20MB');

	PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

	// create new Excel5 reader object for original file 
	$objPHPExcel = PHPExcel_IOFactory::createReader('Excel5');

	// If you only need to access data in your worksheets, and don't need access to the cell formatting, then you can disable reading the formatting information from the workbook:
	//$objPHPExcel->setReadDataOnly(True);

	// can also limit load to specific worksheets to save memory if necessary
	//$objPHPExcel->setLoadSheetsOnly("Sheet1");


	// load data from original excel file
	$objPHPExcel = $objPHPExcel->load($excelFile); 

	// write the data out to an Excel2007 object
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	
	// actual file creation
	$excelFileOut = str_replace(".xls",".xlsx",$excelFile);
	$objWriter->save($excelFileOut);

	// returns new file path
	return $excelFileOut;
}
?>

