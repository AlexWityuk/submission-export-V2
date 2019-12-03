<?php
require_once( AKISMET__PLUGIN_DIR . 'Classes/PHPExcel.php');

$objPHPExcel = new PHPExcel();

$objPHPExcel->getProperties()->setCreator("Tecnavia")
							 ->setLastModifiedBy("Tecnavia")
							 ->setTitle("Office 2007 XLSX Test Document")
							 ->setSubject("Office 2007 XLSX Test Document")
							 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
							 ->setKeywords("office 2007 openxml php")
							 ->setCategory("Submission result file");
$title_cat_name_0= str_replace('&', 'and', $title_cat_name);
/************************************************/
$row = array();
$i = 0;
foreach ($array['items'] as $item) {
	$studio_name = '';
	$referente_name = '';
	$email = '';
	$telephone = '';
	$st = '';
	$pr = '';
	foreach ($item->fields as $field) {
		if ($field['label'] == "Nome studio (si prega di inserire il nome senza farlo precedere da “studio legale”):" ) {
			$studio_name = $field['value'];
		}
		else if ($field['label'] == "Nome referente responsabile: ") {
			$referente_name = $field['value'];
		}
		else if ($field['label'] == "E-mail:") {
			$email = $field['value'];
		}
		else if ($field['label'] == "Telefono: ") {
			$telephone = $field['value'];
		}
  	}
	foreach ($item->fields as $field) {
		$fstr = explode( ' ', $field['label'] );
		$fstr_num = explode( '.', $fstr[0]);
        $f_num = $fstr_num[0];
	  	if( $field['label'] == $f_num . '.1 STUDIO DELL’ANNO' ){
			$st = $field['value'];
	  	}
	  	if( $field['label'] == $f_num . '.2 PROFESSIONISTA DELL’ANNO' ){
			$pr = $field['value'];
	  	}
		if ( $f_num == $cat_num && preg_match("/(\d+)/", $field['label']) && strlen($field['value']) > 0 ) {
			$row[$i]["Nome studio (si prega di inserire il nome senza farlo precedere da “studio legale”):"] = $studio_name;
			$row[$i]["Nome referente responsabile: "] = $referente_name;
			$row[$i]["E-mail:"] = $email;
			$row[$i]["Telefono: "] = $telephone;
			$row[$i][$title_cat_name] = $title_cat_name;
		  	if( $field['label'] == $f_num . '.1 STUDIO DELL’ANNO' ){
				$row[$i][$f_num . '.1 STUDIO DELL’ANNO'][] = $st;
		  	}
		  	if( $field['label'] == $f_num . '.2 PROFESSIONISTA DELL’ANNO' ){
				$row[$i][$f_num . '.2 PROFESSIONISTA DELL’ANNO'][] = $pr;
		  	}
			if (preg_match("/(\d+\.\d+\.\d+\s+)/", $field['label'])) {
				$row[$i][$field['label']][] = $field['value'];
			}
		}
	}
	if (count($row[$i]) > 0) {
		$i++;
	}
}
/*var_dump($row);
wp_die();*/
$page = $objPHPExcel->setActiveSheetIndex(0);

$alph = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');

for ($n=0; $n < count($row); $n++) { 
	$i = 0;

	if ($n < 1) {
		foreach ($row[$n] as $key => $value) {
			$page->setCellValue( $alph[$i] . '1', "$key");
			$i++;
		}
	}
	$i = 0;
	foreach ($row[$n] as $k => $r) {
		if (preg_match("/(\d+\.\d+\.\d+\s+)/", $k)) {
			$objPHPExcel->getActiveSheet()->getColumnDimension($alph[$i])->setWidth(100);
		}
		else $objPHPExcel->getActiveSheet()->getColumnDimension($alph[$i])->setWidth(40);
		$j = $n+2;
		if ( $k == "Telefono: " ) {
			$objPHPExcel->getActiveSheet()->setCellValueExplicit( $alph[$i] . $j, "$r", PHPExcel_Cell_DataType::TYPE_STRING );
		}
		else $page->setCellValue( $alph[$i] . $j, "$r");
		if (preg_match("/(\d+)/", $k)) {
			foreach ($r as $value) {
				$page->setCellValue( $alph[$i] . $j, "$value");
				$j++;
			}
		}
		$i++;
	}
}

/************************Export the Excel************************************************/
$objPHPExcel->getActiveSheet()->setTitle('Simple');
// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);
// Redirect output to a client’s web browser (Excel2007)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="'.$title_cat_name_0 .'.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');

exit();
?>

