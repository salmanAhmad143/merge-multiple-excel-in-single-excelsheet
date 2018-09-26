# Project Title

Merge multiple excel files in single excel sheet using PHP

## Getting Started

This code is help you to merge `n` number of excel files data in a single sheet with in a second. You just need to download the source code and follow the below instructions.


### Prerequisites

This code is using php excel library and offcouse php. So you have to include PHPEXCEL library in you file.

```
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';
require_once 'Classes/PHPExcel/IOFactory.php';

```

### Installing

In this sample code you can see a folder name is "files". There are a number of excel files in it and a file "merge_excel_data_in_single_sheet.php". 

This is the main file which will be use to merge all these files data into the single excel sheet.

```
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

```

## Running the tests

After downloading all the files on your computer you need to run the above file using command prompt on windows system like below :

```

C:\Users\LC-121212\Desktop\merge-excel> php merge_excel_data_in_single_sheet.php C:\Users\LC-121212\Desktop\merge-excel\files

```

### Explain :

1- Go to the directory where put your downloaded folder. In the above example i put it on the desktop. 

2- So first i go into this folder and then i run the "merge_excel_data_in_single_sheet.php" file by using above command.

3- Now after finish the process you will see the data is now stored of all the excel file from "files" folder into "E:\\mergeFileData.xlsx" file.


## Built With

* [PHPEXCEL](https://github.com/PHPOffice/PHPExcel) - The PHP EXCEL library is used.
* [PHP WAMP SERVER](http://www.wampserver.com/en/) - PHP WAMP server is used.

## Contributing

We welcome the new commit of changes in this code. If any body want to contribute in it. (http://phpsollutions.blogspot.com) Please submit a pull requests to us.

## Authors

* **Salman Ahmad** - *Initial work* - [PHPSOLLUTIONS.BLOGSPOT.COM](https://phpsollutions.blogspot.com/p/blog-page.html)

## License

This project is developed using the free open source. So any body are free to download and use this code. 
