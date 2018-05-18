<?php

namespace Gsoft\ExcelBundle;

use Symfony\Component\HttpKernel\Bundle\Bundle;

class GsoftExcelBundle extends Bundle
{
	const CONDITION_CELLIS = PHPExcel_Style_Conditional::CONDITION_CELLIS;
    const OPERATOR_GREATERTHAN = PHPExcel_Style_Conditional::OPERATOR_GREATERTHAN;
    const OPERATOR_LESSTHAN = PHPExcel_Style_Conditional::OPERATOR_LESSTHAN;


    private $sheet;
    private $objPHPExcel;
    private $letters = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
	//private $letters array('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25');
    private $iRow;
    private $iColumn;
    private $colorCell;
    private $colorFont;
    private $nameFont;
    private $sizeFont;
    private $bold;
    private $hascenter;
    private $verticalAlignCenter;

    protected $data;
    protected $tituloArchivo;
    protected $tituloColums;
    protected $hasBorder;
    protected $conditionals;

    public function __construct($data, $tituloColums = null, $tituloArchivo = null, $routeExist = null) {
        $this->data = $data;
        $this->tituloColums = $tituloColums;
        $this->tituloArchivo = ($tituloArchivo == null) ? "Reporte - " : $tituloArchivo;
        $this->init();

        if(is_null($routeExist)){
            $this->createXls();
            $this->setData();
            $this->getExcel();
        }else{
            $this->createFromFile($routeExist);
        }
    }


    /**
     * @Route("/")
     */
    public function indexAction()
    {
        return $this->render('GsoftExcelBundle:Default:index.html.twig');
    }

    protected function createXls() {
        error_reporting(E_ALL);
        ini_set('display_errors', TRUE);
        ini_set('display_startup_errors', TRUE);

        $this->objPHPExcel = new PHPExcel();

        $this->objPHPExcel->getProperties()->setCreator("ExtranetLatinTrade.com")
                ->setLastModifiedBy("ExtranetLatinTrade.com")
                ->setTitle($this->tituloArchivo)
                ->setSubject($this->tituloArchivo)
                ->setDescription($this->tituloArchivo);
    }

    protected function createFromFile($route) {
        try {
            $inputFileType = PHPExcel_IOFactory::identify($route);
            $objReader = PHPExcel_IOFactory::createReader($inputFileType);
            $this->inputFileType=$inputFileType;
            $objReader->setIncludeCharts(TRUE); //Allows displaying graphics
            $this->objPHPExcel = $objReader->load($route);
            return $this->objPHPExcel;
        } catch(Exception $e) {
            die('Error loading file "'.pathinfo($route.' || '.$inputFileType,PATHINFO_BASENAME).'": '.$e->getMessage());
        }
    }

    protected function init() {
        $this->iRow = 1;
        $this->iColum = 0;
        $this->colorCell = NULL;
        $this->colorFont = NULL;
        $this->nameFont = NULL;
        $this->sizeFont = NULL;
        $this->hasBorder = true;
        $this->hascenter = false;

        $this->bold = false;
    }

    public function setActiveSheetIndex($index) {
        $this->objPHPExcel->setActiveSheetIndex($index);
        $this->sheet = $this->objPHPExcel->getActiveSheet();
    }

    protected function setData() {
        $this->iRow = 1;
        $this->setActiveSheetIndex(0);

        if ($this->tituloColums != null) {
            $this->setTitleColums();
            $this->iRow ++;
        }

        foreach ($this->data as $value) {
            $this->setDataByArray($value, $this->iRow);
            $this->iRow++;
        }
    }

    protected function setTitleColums() {
        $this->setDataByArray($this->tituloColums, 1);
    }

    protected function setDataByArray($ArryValues, $iRow, $iColumn = NULL) {
        $this->iColumn = $iColumn == NULL ? 0 : $iColumn;
        foreach ($ArryValues as $value) {
                $this->setValueCell($value, $iRow, $this->iColumn);
            $this->iColumn ++;
        }
    }

    protected function setDataByArrayStyle($arryValues, $arrayStyle, $iRow, $iColumn = NULL) {
        $this->iColumn = $iColumn == NULL ? 0 : $iColumn;
        foreach ($arryValues as $value) {
                if(isset($arrayStyle[$this->iColumn])){
                    $this->resetStyles();
                    $this->setStyleByArray($arrayStyle[$this->iColumn]);
                }
                $this->setValueCell($value, $iRow, $this->iColumn);
            $this->iColumn ++;
        }
        $this->resetStyles();
    }

    protected function setValueCell($value, $iRow, $iColumn,$hasStyle=TRUE) {
        $this->sheet->SetCellValue($this->letters[$iColumn] . $iRow, $value);
        if($hasStyle)
            $this->setStyle($iRow, $iColumn);
    }

    public function getExcel() {
            // Redirect output to a client’s web browser (Excel2007)
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $this->tituloArchivo . '.xlsx"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header('Pragma: public'); // HTTP/1.0

            ob_end_clean();


	        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
	        $objWriter->save('php://output');
	        PHPExcel_Calculation::unsetInstance($this->objPHPExcel);
	        exit;
    }

    protected function getIRow() {
        return $this->iRow;
    }

    protected function getIColumn() {
        return $this->iColumn;
    }

    protected function setIRow($iRow) {
        $this->iRow = $iRow;
    }

    protected function setIColumn($iColumn) {
        $this->iColumn = $iColumn;
    }

    protected function setColorCell($colorCell) {
        $this->colorCell = $colorCell;
    }

    protected function getVerticalAlignCenter() {
        return $this->verticalAlignCenter;
    }

    protected function setVerticalAlignCenter($verticalAlignCenter) {
        $this->verticalAlignCenter = $verticalAlignCenter;
    }

    protected function addCountIrow() {
        $this->iRow ++;
        return $this->iRow;
    }

    protected function setStyle($iRow, $iColumn) {
        if ($this->colorCell)
            $this->cellColor($iRow, $iColumn);

        if($this->hasBorder)
            $this->cellBorder($iRow, $iColumn);

        if($this->hascenter)
            $this->cellCenter($iRow, $iColumn);

        if($this->verticalAlignCenter)
            $this->cellVerticalAlignCenter($iRow, $iColumn);

		//  if($this->conditionals)
		//		$this->setConditionalStyle($iRow, $iColumn);

        $this->setFont($iRow, $iColumn);


    }

    protected function setGeneralFormatNumber($coordenadas) {
        $this->objPHPExcel->getActiveSheet()->getStyle($coordenadas)
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_GENERAL);
    }

    protected function setPercentageFormatNumber($coordenadas) {
        $this->objPHPExcel->getActiveSheet()->getStyle($coordenadas)
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);
        $this->objPHPExcel->getActiveSheet()->getStyle($coordenadas)
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
    }

    protected function setDecimalFormatNumber($coordenadas) {
        $this->objPHPExcel->getActiveSheet()->getStyle($coordenadas)
            ->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00);
        $this->objPHPExcel->getActiveSheet()->getStyle($coordenadas)
            ->getAlignment()
            ->setIndent(0)
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY);
        $this->objPHPExcel->getActiveSheet()->getStyle($coordenadas)
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
    }

    public function getLineHighestSheet(){
        $lastRow = $this->sheet->getHighestDataRow();
        $valor = NULL;
        $count = 0;
        for($i = $lastRow; $i>=0; $i--){
            $valor = $this->getValue($this->letters[0], $i);
            $count++;
        }

        return $count+1;
    }

    protected function getValue($col, $row){
        return $this->sheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
    }

    protected function getColorFont() {
        return $this->colorFont;
    }

    protected function getNameFont() {
        return $this->nameFont;
    }

    protected function getSizeFont() {
        return $this->sizeFont;
    }

    protected function setColorFont($colorFont) {
        $this->colorFont = $colorFont;
    }

    protected function setNameFont($nameFont) {
        $this->nameFont = $nameFont;
    }

    protected function setSizeFont($sizeFont) {
        $this->sizeFont = $sizeFont;
    }

    protected function getBold() {
        return $this->bold;
    }

    protected function setBold($bold) {
        $this->bold = $bold;
    }

    protected function getCell($col,$row){
        return $this->sheet->getCell($col.$row);
    }

    protected function cellColor($iRow, $iColumn, $color = NULL) {
        $color = $color == NULL ? $this->colorCell : $color;

        $this->sheet->getStyle($this->letters[$iColumn] . $iRow)->getFill()->applyFromArray(
            array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => array(
                    'rgb' => $color
                )
            )
        );
    }

    protected function setFont($iRow, $iColumn) {
        $font = $this->colorFont != NULL && $this->nameFont != NULL && $this->sizeFont != NULL;

        if ($font) {
            $styleArray = array(
                'font' => array(
                    'bold' => $this->bold,
                    'color' => array('rgb' => $this->colorFont),
                    'size' => $this->sizeFont,
                    'name' => $this->nameFont
            ));

            $this->sheet->getStyle($this->letters[$iColumn] . $iRow)->applyFromArray($styleArray);
        }
    }

    protected function mergeCell($iRow, $iColumn, $iRowf, $iColumnf) {
		// echo$this->letters[$iColumn] . $iRow . ':' . $this->letters[$iColumnf] . $iRowf;
        $this->sheet->mergeCells($this->letters[$iColumn] . $iRow . ':' . $this->letters[$iColumnf] . $iRowf);
    }

    protected function cellBorder($iRow, $iColumn, $iRowf = NULL, $iColumnf = NULL) {
        $colum = ($iRowf == NULL && $iColumnf == NULL)? $this->letters[$iColumn] . $iRow : $this->letters[$iColumn] . $iRow. ':' . $this->letters[$iColumnf] . $iRowf ;

        $border_style = array(
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('argb' => '000000')
                )
            )
        );

        $this->sheet->getStyle($colum)->applyFromArray($border_style);
    }

    protected function cellCenter($iRow, $iColumn, $iRowf = NULL, $iColumnf = NULL) {
        $colum = ($iRowf == NULL && $iColumnf == NULL)? $this->letters[$iColumn] . $iRow : $this->letters[$iColumn] . $iRow. ':' . $this->letters[$iColumnf] . $iRowf ;

        $style = array(
            'alignment' => array(
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            )
        );

        $this->sheet->getStyle($colum)->applyFromArray($style);
    }

    private function createConditional($firstCodition,$secondCondition,$valueCondition,$color){
        $objConditional = new PHPExcel_Style_Conditional();
        $objConditional->setConditionType($firstCodition)
                        ->setOperatorType($secondCondition)
                        ->addCondition($valueCondition);
        $objConditional->getStyle()->getFont()->getColor()->setARGB($color);
        return $objConditional;
    }

    protected function setConditionalStyle($iRow, $iColumn){
        $conditionalsArray=array();
        foreach($this->conditionals as $conditional){
            $conditionalsArray[] = $this->createConditional($conditional['firstCodition'], $conditional['secondCondition'], $conditional['valueCondition'], $conditional['color']);
        }
        $conditionalStyles = $this->sheet->getStyle($iRow, $iColumn)->getConditionalStyles();
        foreach ($conditionalsArray as $conditional)
            array_push($conditionalStyles, $conditional);
        $this->sheet->getStyle($iRow, $iColumn)->setConditionalStyles($conditionalStyles);
        $this->resetStyles();
    }

    protected function cellVerticalAlignCenter($iRow, $iColumn, $iRowf = NULL, $iColumnf = NULL) {
        $colum = ($iRowf == NULL && $iColumnf == NULL)? $this->letters[$iColumn] . $iRow : $this->letters[$iColumn] . $iRow. ':' . $this->letters[$iColumnf] . $iRowf ;

        $style = array(
            'alignment' => array(
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            )
        );

        $this->sheet->getStyle($colum)->applyFromArray($style);
    }

    protected function createSheet($index){
        $this->objPHPExcel->createSheet($index);
        $this->sheet = $this->objPHPExcel->setActiveSheetIndex($index);
    }

    protected function copyFormatCell($from, $to) {
        $this->sheet->duplicateStyle($this->sheet->getStyle($from), $to);
        $this->sheet->duplicateConditionalStyle($this->sheet->getConditionalStyles($from), $to);
    }

    protected function setStyleByArray($style){
        if(isset($style['colorCell']))
            $this->setColorCell($style['colorCell']);
        if(isset($style['colorFont']))
            $this->setColorFont($style['colorFont']);
        if(isset($style['border']))
            $this->setHasBorder($style['border']);
        if(isset($style['center']))
            $this->setHascenter($style['center']);
        if(isset($style['bold']))
            $this->setBold($style['bold']);
        if(isset($style['sizeFont']))
            $this->setSizeFont($style['sizeFont']);
        if(isset($style['verticalAlignCenter']))
            $this->setVerticalAlignCenter($style['verticalAlignCenter']);
        if(isset($style['conditionals'])){
            $this->setConditionals($style['conditionals']);
        }
    }

    protected function resetStyles(){
        $this->setColorCell(NULL);
        $this->setColorFont(NULL);
        $this->setHasBorder(FALSE);
        $this->setHascenter(FALSE);
        $this->setHascenter(FALSE);
        $this->setConditionals(FALSE);
    }

    public function getDataFromFile($includeHeader = FALSE){
        $highestRow = $this->sheet->getHighestRow();
        $highestColumn = $this->sheet->getHighestColumn();

        $rowBegin=($includeHeader)?1:2;
        $rows=array();
        for ($row = $rowBegin; $row <= $highestRow; $row++){
            $value= $this->sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row,
                                            NULL,
                                            TRUE,
                                            FALSE);
            $value[0]['idRow']=$row;
            $value[0]['is_ok']=1;
            if($value[0][0])
            $rows[$row] = $value[0];
        }
        return $rows;
    }

    protected function setAutoWidthColumn($letterColumn){
        $this->sheet->getColumnDimension($letterColumn)->setAutoSize(true);
    }

    public function getLetters() {
        return $this->letters;
    }

    public function getConditionals() {
        return $this->conditionals;
    }

    public function setConditionals($conditionals) {
        $this->conditionals = $conditionals;
    }

    public function getHasBorder() {
        return $this->hasBorder;
    }

    public function getHascenter() {
        return $this->hascenter;
    }

    public function setHasBorder($hasBorder) {
        $this->hasBorder = $hasBorder;
    }

    public function setHascenter($hascenter) {
        $this->hascenter = $hascenter;
    }

    public function changeDateToFormat($cellCoordinates,$formatPhpDate){
        $cell = $this->getCell($this->letters[$cellCoordinates['col']],$cellCoordinates['row']);
        $numericDate = $cell->getValue();
        if(is_numeric($numericDate)) {
            //se agrega el tiempo equivalente a un dia puesto que phpExcel no retornaba correctamente.
            $time = PHPExcel_Shared_Date::ExcelToPHP($numericDate)+ 86400;
			// $date = date("Y-m-d",$time);
            $date = date($formatPhpDate,$time);
            $this->setValueCell($date, $cellCoordinates['row'], $cellCoordinates['col']);
        }
    }

    public function columnToUppercase($iCol,$limit){
        for($iRow=2;$iRow<=$limit;$iRow++){
            $textCell = $this->getValue($iCol,$iRow);
            $this->setValueCell(strtoupper(trim($textCell)), $iRow, $iCol);
        }
    }
}