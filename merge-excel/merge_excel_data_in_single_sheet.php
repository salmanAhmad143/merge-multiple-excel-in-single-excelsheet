<?php
$filePath =  $argv[1];
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';
require_once 'Classes/PHPExcel/IOFactory.php';
if ($dh = opendir($filePath)){
	$outputFile = "E:\\mergeFileData.xlsx"; //Change the output file location.and file name.
	$objReader = PHPExcel_IOFactory::createReaderForFile($outputFile);
	$objReader->setReadDataOnly(true);
	$objPHPExcel1 = $objReader->load($outputFile);
	
	while (($file = readdir($dh)) !== false) {
		$path_parts = pathinfo($file);		
		if ($path_parts['extension'] === 'xlsx') {			
			$singlefile = $path_parts["filename"];	
			echo "Start merging the content from file " . $singlefile . PHP_EOL;
			$workbook_file = $filePath.'\\'. $file;
			$objReader2 = PHPExcel_IOFactory::createReaderForFile($workbook_file);
			$objReader2->setReadDataOnly(true);
			$objPHPExcel2 = $objReader2->load($workbook_file);
			$objExcel2 = $objPHPExcel2->setActiveSheetIndex(0);

			$findEndDataRow2        = $objExcel2->getHighestRow();
			$findEndDataColumn2     = $objExcel2->getHighestColumn();
			$findEndData2 = $findEndDataColumn2 . $findEndDataRow2;

			$data2 = $objExcel2->rangeToArray('A1:' . $findEndData2);

			unset($objExcel2);
			
			$objExcel1 = $objPHPExcel1->setActiveSheetIndex(0);
			$appendStartRow = $objExcel1->getHighestRow() + 1;
			$objExcel1->fromArray($data2, null, 'A' . $appendStartRow);

			unset($data2);
			echo "End merging the content from file " . $singlefile . PHP_EOL;
		}
	}	
	        
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel1, 'Excel2007');
	$objWriter->save($outputFile);
	
	echo "**********************************************************************". PHP_EOL;
	echo "All the contents of excel file are successfully merged in single file." . PHP_EOL;
	echo "**********************************************************************". PHP_EOL;
 
}



?>