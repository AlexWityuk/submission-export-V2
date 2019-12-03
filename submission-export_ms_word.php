<?php
//require 'vendor/autoload.php';
require_once( AKISMET__PLUGIN_DIR . 'vendor/autoload.php');

// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();

/************List Style******************/

// Define styles
$fontExpWordListStyleName = 'myOwnStyle';
$phpWord->addFontStyle($fontExpWordListStyleName, array(
													'name' => 'Times New Roman', 
													'size' => 11, 
													'bold' => false
													));
$fontHeaderStyle = 'headerOwnStyle';
$phpWord->addFontStyle($fontHeaderStyle, array(
													'name' => 'Times New Roman', 
													'size' => 11, 
													'bold' => false,
													'color' => '827E7D'
													));
$paragraphExpListStyleName = 'P-Style';
$phpWord->addParagraphStyle($paragraphExpListStyleName, array(
													'spaceAfter' => 0,
													'lineHeight' => 1.5,
													)
						);

$phpWord->setDefaultParagraphStyle(
									array(
									'spaceAfter' => \PhpOffice\PhpWord\Shared\Converter::pointToTwip(0),
									'spacing' => 120,
									'lineHeight' => 1.15,
									)
);

/**************************************/
$fontTitle = 'fontTitle';
$phpWord->addFontStyle($fontTitle, array(
											'name' => 'Times New Roman', 
	                    					'size' => 11 ,
	                    					'bold' => false
										));
$paragraphTitle = 'paragraphYitle';
$phpWord->addParagraphStyle($paragraphTitle, array(
							'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER,
	                    	'spaceAfter' => \PhpOffice\PhpWord\Shared\Converter::pixelToTwip(0)
	                    	));

$stodios_array = array();
global $stodios_array;
$subcategory_array = array();
global $subcategory_array;
$sectionName = '***';
foreach ($category_name_array as $title_cat_name => $cat_num) {
	$stodios = array();
	$studio_name_arr_test = array();
	/*------------------------------------------------------------------*/
	foreach ($array['items'] as $item) {
		$i_num = 0;
		$studio_name = '';
		$studio_name_cat_name = false;
		//$studio_name_arr_test = array();
		$last_name_candid = null;
		$first_name_candid = null;
		foreach ($item->fields as $field) {
		/*--------------------------------------------------------------*/
		/*          Get Section Name
		/*--------------------------------------------------------------*/
			if ($field['type'] == 'checkbox' && empty($field['attributes'])) {
				$secNum = explode( '.', $field['label'] );
				if ($secNum[0] == $cat_num) {
					$sectionName = explode( ' ', $field['label']);
					array_shift($sectionName);
					$sectionName = implode(' ', $sectionName);
				}
			}
		/*--------------------------------------------------------------*/
			$field['attributes']['main-category-item']  = str_replace("-", " ", $field['attributes']['main-category-item']);
			if (is_array($field['attributes']) && array_key_exists('cognome-del', $field['attributes']) && strcasecmp($field['attributes']['main-category-item'], $title_cat_name) == 0) {
				$fstr = explode( ' ', $field['label'] );
				$fn = explode( '.', $fstr[0] );
				if ($fn[0] == $cat_num) {
				    $last_name_candid[] = $field;
				}
			}
			if (is_array($field['attributes']) && array_key_exists('nome-del', $field['attributes']) && strcasecmp($field['attributes']['main-category-item'], $title_cat_name) == 0) {
				$fstr = explode( ' ', $field['label'] );
				$fn = explode( '.', $fstr[0] );
				if ($fn[0] == $cat_num) {
	    		    $first_name_candid[] = $field;
				}
			}

			if (isset($field['attributes']['subcategory'])) {
				$subcategory_array[$title_cat_name][] = mb_strtoupper($field['attributes']['subcategory']);
			}
		}
		foreach ($item->fields as $field) {
			if (isset($field['attributes']['subcategory'])) {
				$field['attributes']['main-category-item']  = str_replace("-", " ", $field['attributes']['main-category-item']);
			}
			if (strcasecmp($field['attributes']['main-category-item'], $title_cat_name) == 0) {
				$studio_name_cat_name = true;
			}
		}

		foreach ($item->fields as $field) {
			//----------------------------------------------------------------------------/
			if (is_array($field['attributes']) && array_key_exists('studio-name', $field['attributes']) && $studio_name_cat_name ) {
			    $i_num = 0;
			    foreach ($studio_name_arr_test as $val) {
			    	if (strcasecmp($val, $field['value']) == 0) {
			    		$i_num++;
			    	}
			    }
			    if ($i_num == 0) {
			    	$studio_name = $field['value'];
			    }
			    else $studio_name = $field['value'].' (duplicato '. $i_num.')';
			    $studio_name_arr_test[] = $studio_name;
			    //$studio_name = $field['value'];
			    /*
			    //var_dump($studio_name);
			    //echo "</br>";
			    */
			}
			//------------------------------------------------------------------------------/
			$sub_arr = array();
			$field_arr = array();
			/*------------------------------------------------*/
			if (isset($field['attributes']['subcategory'])) {
				$field['attributes']['main-category-item']  = str_replace("-", " ", $field['attributes']['main-category-item']);
			}
			/*------------------------------------------------*/
			$fstr = explode( ' ', $field['label'] );
	        $f_num = $fstr[0];
	        /*--------------------------------------------------------------*/
	        $f_num = explode( '.', $f_num);
	        $f_num = $f_num[0];
	        /*--------------------------------------------------------------*/
	        $mandato = $fstr[1] . ' ' . $fstr[2];
	        $lab_arr = array();
	        if (preg_match("/(\d+\:\s+)/", $field['label'])) {
	        	$lab_arr =  preg_split("/(\d+\:\s+)/", $field['label']);
	        }
	       
	        else {
	        	$lab_arr =  preg_split("/(\.\d+\s+)/", $field['label']);
	        	
	        }
	        $label  = str_replace("(max 500 battute)","", $lab_arr[1] );
	        $label  = str_replace(", max 500 battute","",  $label );
	        $label  = str_replace(", max 750 battute","",  $label );
	        if ($f_num == $cat_num && $field['type'] == "textarea" && strlen($field['value']) > 0) {
		        $field['value'] = trim($field['value']);
	        }
	        if ($f_num == $cat_num /*&& $field['type'] == "textarea"*/ && array_key_exists('show', $field['attributes']) && $field['attributes']['show'] == 'yes' && strcasecmp($field['attributes']['main-category-item'], $title_cat_name) == 0 && strlen($field['value']) > 0) {
	        	//$itemid[] = $item->id;
	            /*------------------------------*/
	            	    /*var_dump($studio_name);
			    echo "</br>";*/
			    
	            $studio_name = str_replace(" ","_",$studio_name);//////////////////////
				$st_name_arr = explode('&', $studio_name);
				//$studio_name = implode('', $st_name_arr);
				$studio_name = implode('**', $st_name_arr);
				$studio_name = str_replace('__', "_**_", $studio_name);///
				$studio_name = str_replace('[]', "[**]", $studio_name);
				//$studio_name = str_replace('amp;', "**", $studio_name);
				$studio_name = str_replace('amp;', "", $studio_name);
				$studio_name = mb_strtoupper($studio_name);
				$studio_name = str_replace('**', "&", $studio_name);
				$studio_name = str_replace(",","*",$studio_name);//////////////////////
				$studio_name = str_replace("*",",",$studio_name);
				
				//$studio_name = htmlspecialchars ( $studio_name );
				//$studio_name = str_replace('#039;', "'", $studio_name);
				//$studio_name = str_replace('QUOT;', '"', $studio_name);
	        	/*-------------------------------*/
				$subcat_arr = explode('-', $field['attributes']['subcategory']);
				$subcategory = implode(' ', $subcat_arr);
				$subcategory = mb_strtoupper($field['attributes']['subcategory']/*$subcategory*/);
				$subcatit_arr = explode('-', $field['attributes']['subcategory-item']);
				$subcategory_item = implode(' ', $subcatit_arr);
				$subcategory_item = $subcategory_item .':';
				if ($first_name_candid != null) {
					/*------------------------------------------------------------------------------------*/
					for ($i=0; $i < count( $first_name_candid) ; $i++) { 
						$last_name_candid[$i]['value'] = mb_strtoupper($last_name_candid[$i]['value']);
						$first_name_candid[$i]['value'] = mb_strtoupper($first_name_candid[$i]['value']);
						$last_name_candid[$i]['value'] = trim($last_name_candid[$i]['value']);
						$first_name_candid[$i]['value'] = trim($first_name_candid[$i]['value']);
						if ($field['attributes']['subcategory'] != 'no' && $last_name_candid[$i]['attributes']['subcategory'] != 'no') {
							if ($field['attributes']['subcategory'] == $last_name_candid[$i]['attributes']['subcategory'] || $field['attributes']['subcategory'] == $first_name_candid[$i]['attributes']['subcategory'] ) {
								$stodios[$last_name_candid[$i]['value']  . 
					         	' ' . $first_name_candid[$i]['value'] .'$' . $studio_name ][$subcategory][$subcategory_item][] =  array(
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
						else {
							if ($field['attributes']['subcategory-item'] == $last_name_candid[$i]['attributes']['subcategory-item'] || $field['attributes']['subcategory-item'] == $first_name_candid[$i]['attributes']['subcategory-item'] ) {
								$stodios[$last_name_candid[$i]['value']  . 
					         	' ' . $first_name_candid[$i]['value'] .'$' . $studio_name ][$subcategory][$subcategory_item][] =  array(
																									'label-field' => $label,
																									'value-field' => $field['value'],
																									'attributes' => $field['attributes']
																								);
							}
							else if (!array_key_exists('subcategory-item', $last_name_candid[$i]['attributes'])) {
								$stodios[$last_name_candid[$i]['value']  . 
					         	' ' . $first_name_candid[$i]['value'] .'$' . $studio_name ][$subcategory][$subcategory_item][] =  array(
																									'label-field' => $label,
																									'value-field' => $field['value'],
																									'attributes' => $field['attributes']
																								);
							}
						}
					}
					/*----------------------------------------------------------------------------------------*/
				}
				else {
		            /*var_dump($studio_name);
		            echo strlen($studio_name);
	            	echo "</br>";*/
					$arr = array(
							'label-field' => $label,
							'value-field' => $field['value'],
							'attributes' => $field['attributes']
									);
					$stodios[$studio_name][$subcategory][$subcategory_item][] =  $arr;
				}
			}
		}
	}
	$stodios_array[$title_cat_name] = $stodios;

	$subcategory_array[$title_cat_name] = array_unique($subcategory_array[$title_cat_name]);

	/*------------------------------------------------------------------*/
}


function makeParagraph( $section, $title_cat_name, $key, $subcatname, $subcategory_item ) {

	global $stodios_array;
	$phpWord = new \PhpOffice\PhpWord\PhpWord();
	$i = 0;
	$mandatonum = $subcategory_item;
	$stodios = $stodios_array[$title_cat_name];

	if (isset($stodios[$key][$subcatname][$mandatonum])) {
		$arr_reverse = $stodios[$key][$subcatname][$mandatonum];
		$i = 0;
		foreach ($arr_reverse as $value) {
			$new_arr_title = array();
			$new_arr_content = array();
			$result_content = array('name' => 'Times New Roman', 'size' => 11);
			$result_title = array(
					              	'italic' => true,
					              	'name' => 'Times New Roman', 
									'size' => 11
					        	);
			$key_ = str_replace('&', "8", $key);
			$key_ = str_replace('amp;', "", $key_);
			$fontTextStyleName_title = 'fontstyle title**'.$key_.'**'.$i.'**'.$mandatonum;
			$fontTextStyleName_content = 'fontstyle content**'.$key_.'**'.$i.'**'.$mandatonum;
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
			$value['label-field'] = ucfirst($value['label-field']); //with Capital letter 
			$value['label-field'] = trim($value['label-field']);
			if (strlen($subcategory_item)!=0 && $i == 0) {
				$section->addText($subcategory_item ." ". $value['label-field'], $fontTextStyleName_title);
			}
			else {
				$section->addText($value['label-field'], $fontTextStyleName_title);
			}
	     	$value['value-field'] = strip_tags($value['value-field']);
	     	$value['value-field'] = htmlspecialchars ( $value['value-field']);
			$section->addText($value['value-field'], $fontTextStyleName_content);
			//var_dump($value['value-field']);
			$section->addTextBreak(1);
			$i++;
		}
	}
}


$title_for_multiple_exp;
foreach ($stodios_array as $title_cat_name => $stodios) {
	$title_cat_name_0 = str_replace('&', 'and', $title_cat_name);
	$title_for_multiple_exp = $title_for_multiple_exp.$title_cat_name_0.' ';
}
$title_for_multiple_exp = trim($title_for_multiple_exp);
$n = 1;
foreach ($stodios_array as $title_cat_name => $stodios) {
	$title_cat_name_0 = str_replace('&', 'and', $title_cat_name);
	$itritr = 0;
	if (count($stodios) > 0) {
		//------------------------------------------------------------------------/
		$section = $phpWord->addSection( array(
							        'marginLeft'   => 1134.328,
							        'marginRight'  => 1134.328,
							        //'marginTop'    => 1417.91,
							        'marginTop'    => 1000.91,
							        'marginBottom' => 1134.328,
							        'headerHeight' => 30,
							        'footerHeight' => 50,
							    )
							);
		//------------------------------------------------------------------------/
		/*-----------------------------------------------------------*/
		/*    Add header
		/*-----------------------------------------------------------*/
		if ($id == 17) {
			// Add header
			$header = $section->createHeader();
			$table = $header->addTable();
			$headerStr = 'TopLegal Awards – '. $sectionName;
			$table->addRow();
			$table->addCell(10500)->addText('');
			$table->addRow();
			$table->addCell(11000)->addText($headerStr);

			$table->addCell(3500)->addText($title_for_multiple_exp);
		} 
		/*-----------------------------------------------------------*/
		$listFormat = $phpWord->addNumberingStyle(
		            'multilevel-1'.$n,
		            array('type' => 'multilevel', 'levels' => array(
		                    array('format' => 'decimal', 'text' => '%1.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720)                    
		                )
		            )
		        );
		$listFormat = $phpWord->addNumberingStyle(
		            'multilevel-2'.$n,
		            array('type' => 'multilevel', 'levels' => array(
		                    array('format' => 'decimal', 'text' => '%1.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720)                    
		                )
		            )
		        );
		//------------------------------------------------------------------------/
		/*-------------------------------------------------------------------------*/
		ksort($stodios);
		$parsize = 0;
		foreach ($subcategory_array[$title_cat_name] as $subcat_name) {
			$parsize++;
			//var_dump($subcat_name);
			$suname_short_arr = explode(' ', $subcat_name);
			$subname_short = $suname_short_arr[0];
			$i = 0;
			foreach ($stodios as $key => $subcats) {
				foreach ($subcats as $kk => $value) {
					if ($kk == $subcat_name) {
						//var_dump($value);
						if ($i == 0) {
							if ($itritr > 0) $section->addTextBreak(3);
							$itritr++;
							$title_cat_name_1 = '';
							$title_cat_name_1 = mb_strtoupper($title_cat_name_0);
							$subname_short = explode('-', $subname_short);
							$subname_short = implode(' ', $subname_short);
							/* //Old aspect
							if (strcasecmp('no', $subname_short) == 0) {
								$title_cat_name_1 = mb_strtoupper($title_cat_name_0);
							}
							else {
								$subname_short = explode('-', $subname_short);
								$subname_short = implode(' ', $subname_short);
								$title_cat_name_1 = strtoupper($title_cat_name_0 . " (". $subname_short. ")");
							}
							*/
							$kf = ceil(strlen($title_cat_name_1)/47);
							$koeff = 38 * $kf;
							$textbox = $section->addTextBox(
							    array(
							        'alignment'   => \PhpOffice\PhpWord\SimpleType\Jc::CENTER,
							        'width'       => 650,
							        'height'      => 180,
							        'borderSize'  => 0,
							        'borderColor' => '#fff'
							        //'borderColor' => '#1b98e0'
							    )
							);
							$textbox->addText($title_cat_name_1, $fontTitle, $paragraphTitle);
							if ( $subname_short != 'NO' ) {
								$textbox->addText($subname_short, $fontTitle, $paragraphTitle);
							}
							/*
							$section->addText( $title_cat_name_1, $fontTitle, $paragraphTitle );
							$section->addText( $subname_short, $fontTitle, $paragraphTitle );
							*/
							$section->addTextBreak(1);
							//echo $title_cat_name_1."</br>";
						}
						if( !preg_match("/([\$]+)/", $key) ) { 
							/*--------------------------*/
							$key = explode('_', $key);
							$key = implode(' ', $key);
							$key = str_replace('&#039;', "'", $key);
							$key = htmlspecialchars ( $key );
							//$key = str_replace('&#039;', "'", $key);
							$key = str_replace('#039;', "'", $key);
							$key = str_replace('QUOT;', '"', $key);
							/*--------------------------*/
					        $section->addListItem($key, 0, $fontExpWordListStyleName, 'multilevel-1'.$n, $paragraphExpListStyleName);
					  		//echo $key."</br>";
					  	}
					  	else {
					  		$arr = explode( '$', $key );
							$ncomp = $arr[1];
							/*------------------------------*/
							$ncomp_ = explode('_', $ncomp);
							$ncomp = implode(' ', $ncomp_);
							$ncomp = htmlspecialchars ( $ncomp );
							$ncomp = str_replace('#039;', "'", $ncomp);
							$ncomp = str_replace('QUOT;', '"', $ncomp);
							/*------------------------------*/
							$name = $arr[0];
					        $section->addListItem($name . ' (' . $ncomp . ')', 0, $fontExpWordListStyleName, 'multilevel-2'.$n, $paragraphExpListStyleName);
					  		//echo "+++".$key."</br>";
					  	}
			  			$i++;
					}
				}/*foreach ($subcats as $kk => $value)*/
			}
		  	/*$section->addTextBreak(2);*/
		  	$section->addPageBreak();
		  	$nn = 0;
		  	$table = $section->addTable();
		  	if ( $subname_short != 'NO' ) {
		  		$headerStr = $title_cat_name_1.' – '. $subname_short . ' ' . date("Y");
		  	} else
			$headerStr = $title_cat_name_1.' – '. date("Y");
			$table->addRow();
			$table->addCell(15000);
			$table->addCell(11000)->addText($headerStr, $fontHeaderStyle);
			foreach ($stodios as $key => $subcats) {
				foreach ($subcats as $kk => $value) {
					if ($kk == $subcat_name) {
						if ($nn == 0) $section->addTextBreak(2);
						if( !preg_match("/([\$]+)/", $key) ) {
							//---------------------------/
							$key_ = explode('_', $key);
							$key_ = implode(' ', $key_);
							$key_ = htmlspecialchars ( $key_ );
							$key_ = str_replace('#039;', "'", $key_);
							$key_ = str_replace('QUOT;', '"', $key_);
							//---------------------------/
						  	$section->addText($key_, $fontExpWordListStyleName);
						} 
						else {
							$arr = explode( '$', $key );
							$ncomp = $arr[1];
							//------------------------------/
							$ncomp_ = explode('_', $ncomp);
							$ncomp = implode(' ', $ncomp_);
							$ncomp = htmlspecialchars ( $ncomp );
							$ncomp = str_replace('#039;', "'", $ncomp);
							$ncomp = str_replace('QUOT;', '"', $ncomp);
							//------------------------------/
							$name = $arr[0];
							$section->addText($name . ' (' . $ncomp . ')', $fontExpWordListStyleName);
						}
						//$section->addTextBreak(1);
						//-------------------------------------------------------/
						foreach ($subcats as $kk => $subcategory_item) {
							foreach ($subcategory_item as $k => $value) {
								makeParagraph( $section, $title_cat_name, $key, $kk, $k);
								//var_dump($value);
							}
							//echo "</br>";
						}
						//$section->addTextBreak(1);
						$nn++;
					}
				}//foreach ($subcats as $kk => $value)
				$section->addTextBreak(6);
			} 
			
			if ( $parsize < count( $subcategory_array[$title_cat_name] ) ) {
				$section->addPageBreak();
			} 
			//$section->addPageBreak();
			//if ($itritr < count($subcategory_array[$title_cat_name])) $section->addTextBreak(2);
		} 
	}
	$n++;	
	/*-------------------------------------------------------------------------*/
}
$title_for_multiple_exp;
foreach ($stodios_array as $title_cat_name => $stodios) {
	$title_cat_name_0 = str_replace('&', 'and', $title_cat_name);
	$title_for_multiple_exp = $title_for_multiple_exp.$title_cat_name_0.' ';
}
$title_for_multiple_exp = trim($title_for_multiple_exp);
//wp_die();
//var_dump($title_cat_name_0);
//wp_die();
/************************************/
/************************Export the Word************************************************/
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$file_name = 'submissionxport.docx';
$objWriter->save(AKISMET__PLUGIN_DIR . $file_name);
//var_dump(AKISMET__PLUGIN_DIR . $file_name);
//wp_die();
/*header("Content-type: application/vnd.ms-word");
header("Content-Disposition: attachment;Filename=". 'submissionxport.doc');
exit();*/

header('Content-Description: File Transfer');
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="'.$title_for_multiple_exp.'.docx"');
header('Content-Transfer-Encoding: binary');
header('Expires: 0');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Pragma: public');
header('Content-Length: ' . filesize(AKISMET__PLUGIN_DIR . $file_name));
flush();
readfile(AKISMET__PLUGIN_DIR . $file_name);
unlink(AKISMET__PLUGIN_DIR . $file_name); 
exit();
