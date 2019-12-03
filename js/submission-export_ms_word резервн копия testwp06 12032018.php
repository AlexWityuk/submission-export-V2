<?php
//require 'vendor/autoload.php';
require_once( AKISMET__PLUGIN_DIR . 'vendor/autoload.php');

// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();

/************List Style******************/

// Define styles
$fontExpWordListStyleName = 'myOwnStyle';
$phpWord->addFontStyle($fontExpWordListStyleName, array(
													'name' => 'Calibri', 
													'size' => 11, 
													'bold' => true
													));
$paragraphExpListStyleName = 'P-Style';
$phpWord->addParagraphStyle($paragraphExpListStyleName, array('spaceAfter' => 0));

$phpWord->setDefaultParagraphStyle(
									array(
									'spaceAfter' => \PhpOffice\PhpWord\Shared\Converter::pointToTwip(0),
									'spacing' => 120,
									'lineHeight' => 1,
									)
);

$listFormat = $phpWord->addNumberingStyle(
            'multilevel-1',
            array('type' => 'multilevel', 'levels' => array(
                    array('format' => 'decimal', 'text' => '%1.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720)                    
                )
            )
        );
$listFormat = $phpWord->addNumberingStyle(
            'multilevel-2',
            array('type' => 'multilevel', 'levels' => array(
                    array('format' => 'decimal', 'text' => '%1.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720)                    
                )
            )
        );

/**************************************/
$fontTitle = 'fontTitle';
$phpWord->addFontStyle($fontTitle, array(
											'name' => 'Garamond', 
	                    					'size' => 18 ,
	                    					'bold' => true
										));
$paragraphTitle = 'paragraphYitle';
$phpWord->addParagraphStyle($paragraphTitle, array(
							'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER,
	                    	'spaceAfter' => \PhpOffice\PhpWord\Shared\Converter::pixelToTwip(0)
	                    	));

// Add an Section to the document...
$section = $phpWord->addSection( array(
						        'marginLeft'   => 1134.328,
						        'marginRight'  => 1134.328,
						        'marginTop'    => 1417.91,
						        'marginBottom' => 1134.328,
						        'headerHeight' => 50,
						        'footerHeight' => 50,
						    )
						);
$title_cat_name_0 = str_replace('&', 'and', $title_cat_name);

$stodios = array();
global $stodios;
$subcategory_array = array();
global $subcategory_array;
//var_dump($array['items']);
foreach ($array['items'] as $item) {
	$studio_name = '';
	$last_name_candid = null;
	$first_name_candid = null;
	foreach ($item->fields as $field) {
		if (is_array($field['attributes']) && array_key_exists('studio-name', $field['attributes'])) {
		    $studio_name = $field['value'];
		}
		if (is_array($field['attributes']) && array_key_exists('cognome-del', $field['attributes'])) {
		    $last_name_candid = array(
		    						'cognome' => $field['value'] ,
		    						'subcategory' => $field['attributes']['subcategory']
		    						);
		}
		if (is_array($field['attributes']) && array_key_exists('nome-del', $field['attributes'])) {
		    $first_name_candid = array(
		    						'nome' => $field['value'] ,
		    						'subcategory' => $field['attributes']['subcategory']
		    						);
		}
	}
	foreach ($item->fields as $field) {
		$sub_arr = array();
		$field_arr = array();
		/*------------------------------------------------*/
		if (isset($field['attributes']['subcategory'])) {
			$field['attributes']['main-category-item']  = str_replace("-", " ", $field['attributes']['main-category-item']);
			//var_dump($field['attributes']['subcategory']);
			//echo "<br/>".$field['attributes']['main-category-item']."****".$title_cat_name;
		}
		/*------------------------------------------------*/
		$fstr = explode( ' ', $field['label'] );
        $f_num = $fstr[0];
        $mandato = $fstr[1] . ' ' . $fstr[2];
        $lab_arr =  preg_split("/(\d+\:\s+)/", $field['label']);
        $label = explode( '(max 500 battute)', $lab_arr[1] );
        $label = $label [0];
        $label = explode( ', max 500 battute', $label );
        $label = $label [0];
        if ( $field['type'] == "textarea" /*&& $field['attributes']['show'] == 'yes'*/ && $field['attributes']['main-category-item'] === $title_cat_name && strlen($field['value']) > 0) {
            /*------------------------------*/
			$st_name_arr = explode('&', $studio_name);
			$studio_name = implode('', $st_name_arr);
			$studio_name = str_replace('amp;', "**", $studio_name);
			$studio_name = mb_strtoupper($studio_name);
			$studio_name = str_replace('**', "&", $studio_name);
			$studio_name = htmlspecialchars ( $studio_name );
			$studio_name = str_replace('#039;', "'", $studio_name);
        	/*-------------------------------*/
			$subcat_arr = explode('-', $field['attributes']['subcategory']);
			$subcategory = implode(' ', $subcat_arr);
			$subcategory = mb_strtoupper($field['attributes']['subcategory']/*$subcategory*/);
			$subcategory_array[] = mb_strtoupper($field['attributes']['subcategory']);//$subcategory;
			//var_dump($subcategory_array);
			$subcatit_arr = explode('-', $field['attributes']['subcategory-item']);
			$subcategory_item = implode(' ', $subcatit_arr);
			$subcategory_item = $subcategory_item .':';
			if ($last_name_candid['subcategory'] == $field['attributes']['subcategory'] || $first_name_candid['subcategory'] == $field['attributes']['subcategory']) {
				$stodios[$last_name_candid['cognome']  . 
		         	' ' . $first_name_candid['nome'] .'$' . $studio_name ][$subcategory][$subcategory_item][] =  array(
																						'label-field' => $label,
																						'value-field' => $field['value'],
																						'attributes' => $field['attributes']
																					);
			}
			else {
				$stodios[$studio_name][$subcategory][$subcategory_item][] =  array(
					'label-field' => $label,
					'value-field' => $field['value'],
					'attributes' => $field['attributes']
							);
			}
		}
	}
}
//var_dump($subcategory_array);
//wp_die();
$subcategory_array = array_unique($subcategory_array);
//wp_die();
function makeParagraph( $section, $key, $subcatname, $subcategory_item ) {
	global $stodios;
	$phpWord = new \PhpOffice\PhpWord\PhpWord();
	$i = 0;
	$mandatonum = $subcategory_item;
	if (isset($stodios[$key][$subcatname][$mandatonum])) {
		$arr_reverse = $stodios[$key][$subcatname][$mandatonum];
		$i = 0;
		foreach ($arr_reverse as $value) {
			$new_arr_title = array();
			$new_arr_content = array();
			$result_content = array('name' => 'Calibri', 'size' => 11);
			$result_title = array(
					              	'italic' => true,
					              	'name' => 'Calibri', 
									'size' => 11
					        	);
			$fontTextStyleName_title = 'fontstyle title**'.$key.'**'.$i.'**'.$mandatonum;
			$fontTextStyleName_content = 'fontstyle content**'.$key.'**'.$i.'**'.$mandatonum;
			foreach ($value['attributes'] as $key => $val) {
				$param = explode('-', $key);
				$endelem = array_pop($param);
				if ($endelem == 'title') {
					$new_arr_title[$param[1]] = $val;
				}
				else if ($endelem == 'content') {
					$new_arr_content[$param[1]] = $val;
				}
			}
			$result_title = array_merge($result_title, $new_arr_title);
			$result_content = array_merge($result_content, $new_arr_content);
			$phpWord->addFontStyle($fontTextStyleName_title, $result_title);
			$phpWord->addFontStyle($fontTextStyleName_content, $result_content);
			$subcategory_item = str_replace('disabled:', "", $subcategory_item);//////////////////
			if ($i == 0) {
				$section->addText($subcategory_item . $value['label-field'], $fontTextStyleName_title);
			}
			else $section->addText($value['label-field'], $fontTextStyleName_title);
	     	$value['value-field'] = strip_tags($value['value-field']);
	     	$value['value-field'] = htmlspecialchars ( $value['value-field']);
			$section->addText($value['value-field'], $fontTextStyleName_content);
			$section->addTextBreak(1);
			$i++;
		}
	}
}
/*foreach ($stodios as $key => $subcats) {
	echo $key ."//</br>";
	echo "<br/>".key($subcats)."**777***" ;
	foreach ($subcats as $kk => $value) {
		echo $kk ."***</br>";
		foreach ($value as $k => $val) {
			echo $k."</br>";
			var_dump($val);
			echo "</br>".$k;
			echo "</br>";
		}
		echo "</br>***".$kk;
		echo "</br>";
	}
	echo "</br>//".$key;
	echo "</br>";
}
wp_die();*/
ksort($stodios);
foreach ($subcategory_array as $subcat_name) {
	$suname_short_arr = explode(' ', $subcat_name);
	$subname_short = $suname_short_arr[0];
	$i = 0;
	foreach ($stodios as $key => $subcats) {
		foreach ($subcats as $kk => $value) {
			if ($kk == $subcat_name) {
				if ($i == 0) {
					$title_cat_name_1 = strtoupper($title_cat_name_0 . " (". $subname_short. ")");
					$kf = ceil(strlen($title_cat_name_1)/41);
					$koeff = 47 * $kf;
					//echo (int)$koeff."<br/>";
					$textbox = $section->addTextBox(
					    array(
					        'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::CENTER,
					        'width'       => 650,
					        'height'      => $koeff,
					        'borderSize'  => 1,
					        'borderColor' => '#1b98e0'
					    )
					);
					$textbox->addText($title_cat_name_1, $fontTitle, $paragraphTitle);
					$section->addTextBreak(2);
					//echo $title_cat_name_1."</br>";
				}
				if( !preg_match("/([\$]+)/", $key) ) { 
			        $section->addListItem($key, 0, $fontExpWordListStyleName, 'multilevel-1', $paragraphExpListStyleName);
			  		//echo $key."</br>";
			  	}
			  	else {
			  		$arr = explode( '$', $key );
					$ncomp = $arr[1];
					$name = $arr[0];
			        $section->addListItem($name . ' (' . $ncomp . ')', 0, $fontExpWordListStyleName, 'multilevel-2', $paragraphExpListStyleName);
			  		//echo "+++".$key."</br>";
			  	}
	  			$i++;
			}
		}/*foreach ($subcats as $kk => $value)*/
	}
  	$section->addTextBreak(2);
	foreach ($stodios as $key => $subcats) {
		foreach ($subcats as $kk => $value) {
			if ($kk == $subcat_name) {
				if( !preg_match("/([\$]+)/", $key) ) {
				  	$section->addText($key, $fontExpWordListStyleName);
				} 
				else {
					$arr = explode( '$', $key );
					$ncomp = $arr[1];
					$name = $arr[0];
					$section->addText($name . ' (' . $ncomp . ')', $fontExpWordListStyleName);
				}
				$section->addTextBreak(1);
				/*-------------------------------------------------------*/
				foreach ($subcats as $kk => $subcategory_item) {
					foreach ($subcategory_item as $k => $value) {
						makeParagraph( $section, $key, $kk, $k);
						//var_dump($value);
					}
					//echo "</br>";
				}
				$section->addTextBreak(1);
			}
		}/*foreach ($subcats as $kk => $value)*/
	}
}

//wp_die();
/************************************/
/************************Export the Word************************************************/
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$file_name = 'submissionxport.docx';
$objWriter->save(AKISMET__PLUGIN_DIR . $file_name);

/*header("Content-type: application/vnd.ms-word");
header("Content-Disposition: attachment;Filename=". 'submissionxport.doc');
exit();*/

header('Content-Description: File Transfer');
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="'.$title_cat_name_0.'.docx"');
header('Content-Transfer-Encoding: binary');
header('Expires: 0');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Pragma: public');
header('Content-Length: ' . filesize(AKISMET__PLUGIN_DIR . $file_name));
flush();
readfile(AKISMET__PLUGIN_DIR . $file_name);
unlink(AKISMET__PLUGIN_DIR . $file_name); 
exit();
