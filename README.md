# gsExcel
makes phpexcel simple and usable.

````PHP
//Returns a simple .xlsx file with three columns 
	$arrayRows(	
			('Lucia','Cll 80b # 45 - 6', '6453827'),
			('George','Cll 95b # 67 - 6', '85847635')
		);
	$titles = array('Name','Address','Phone');
return new GsExcel($arrayRows,$titles,'Users-Report');
	
// OR Create on base to other file
$file = new GsExcel($arrayRows,$titles,'Users-Report');
$file->createFromFile(sfConfig::get('sf_upload_dir').'/Formatos/FormatoReportePlataformas.xlsx');
$file->init();
$file->setTitles($tituloColums,$this->iRow);
$file->getExcel();

````