<?php
/*
Plugin Name: Plugin by exporting the submissions
Plugin URI: 
Description: Plugin by exporting the submissions, of a specific form, in different formats (Word, excel, ecc..) 
Version: 1.0
Author: Alex Wityuk
Author URI:
*/

define( 'AKISMET__PLUGIN_DIR', plugin_dir_path( __FILE__ ) );
require_once( AKISMET__PLUGIN_DIR . 'submission-export_wplist_table.php');
require_once( AKISMET__PLUGIN_DIR . 'vendor/autoload.php');
function submission_export_enqueue_style_admin() 
{
	wp_enqueue_style( 'custom_admin_submission_export_css', plugins_url('css/admin.css', __FILE__));
	wp_enqueue_script( 'submission_export_custom_script', plugins_url('/js/script.js', __FILE__));
}
add_action ( 'admin_enqueue_scripts', 'submission_export_enqueue_style_admin' );

function add_submission_export_page()
{
        add_submenu_page(
            'manage_fm',
            'Exporting the submissions', 
            'Form Maker Exporter',
            'manage_options', 
            'primer_slug', 'submission_export_options_page_output');

}
add_action ('admin_menu', 'add_submission_export_page');

function submission_export_options_page_output ()
{
  	$myListTable = new Export_Forms_List_Table();
    $myListTable->prepare_items();
    ?>
    <div class="wrap">    
        <div id="icon-users" class="icon32"><br/></div>
        <h2>The list of the forms for the submissions export</h2>     
        <form id="submission_export_filter" method="post" action="">
        	<input type="hidden" name="page" value="my_list_test" />
			<?php $myListTable->search_box('search', 'search_id'); ?>
            <input type="hidden" name="page" value="<?php echo $_REQUEST['page'] ?>" />
            <?php $myListTable->display() ?>
            <p style="margin: 30px"></p>
            <input type="submit" class="button-primary" name="convert_to_json" value="<?php _e('To export the JSON') ?>" />
            <input type="submit" class="button-primary" name="convert_to_word" value="<?php _e('To export the MS Word') ?>" />
            <input type="submit" class="button-primary" name="convert_to_excel" value="<?php _e('To export the MS Excel') ?>" />
        </form>
    </div>
    <?php
}
function submission_export_to_json_file() {
    if ( isset($_POST['exptojson']) ) {
        $id = (int)$_POST['exptojson'][0];
        $array = submission_export_Get_Submission($id);
        if ( isset($_POST['convert_to_json']) ) {
            $file_name = 'tablexport.json';
            header('Content-Type: application/json; charset=utf-8');
            echo  json_encode($array, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
            header("Content-disposition: attachment; filename=".$file_name);
            exit();
        }
        if ( isset($_POST[$id.'expformoption']) ) {
            /*---------------------------------------------------------*/
            $category_name_array = array();
            for ($j=0; $j<count($_POST[$id.'expformoption']); $j++) {
                $category_name  = str_replace('\\', '', $_POST[$id.'expformoption'][$j]);
                $astr = explode( ' ', $category_name );
                $cat_num = explode( '.', $astr[0]);
                $cat_num = $cat_num[0]; //number of category
                $title_cat_name = '';
                for ($i=1; $i < count($astr); $i++) { 
                    $title_cat_name = $title_cat_name . $astr[$i] . ' ';
                }
                $title_cat_name = substr($title_cat_name, 0, -1);
                $category_name_array[$title_cat_name] = $cat_num;
            }
            /*----------------------------------------------------------*/
            if ( isset($_POST['convert_to_word']) ) {
                require_once( AKISMET__PLUGIN_DIR . 'submission-export_ms_word.php');
            }
            else if ( isset($_POST['convert_to_excel']) ) {
                require_once( AKISMET__PLUGIN_DIR . 'submission-export_ms_excel.php');
            }
        }
    }
    else return;
}

function submission_export_Get_Submission($id) {
    global $wpdb;
    $table_name = $wpdb->prefix . "formmaker";
    $query = "SELECT title, form_fields FROM $table_name WHERE id = $id";
    $result = $wpdb->get_results($query);
    $arr = explode('*:*new_field*:*', $result[0] -> form_fields );
    /*var_dump($arr);
    wp_die();*/
    $fields = array();
    for ( $i = 0; $i < count($arr); $i++ ) {
        $value = explode( '*:*', $arr[$i] );
        $fields[] = $value;
    }
    /*var_dump($fields);
    wp_die();*/
    for ( $i = 0; $i < count($fields); $i++) {
        $item_obj = array();
        $sub_arr = array();
        for ( $j = 0; $j < count($fields[$i]); ) {
            $key = $fields[$i][$j + 1];
            $value = $fields[$i][$j];
            if ($key == 'w_attr_name') {
                $sub_arr[] = $value;
            }
            else $item_obj[$key] = $value;
            $j+=2;
        }
        $item_obj['w_attr_name'] = $sub_arr;
        $fields[$i] = $item_obj;
        $attributes_arr = array();

        for ($n=0; $n < count($fields[$i]['w_attr_name']); $n++) { 
            $attributes = explode( '=', $fields[$i]['w_attr_name'][$n]);
            $attributes_arr[$attributes[0]] = $attributes[1];
        }
        $fields[$i]['w_attr_name'] = $attributes_arr;
    }
    for ( $i = 0; $i < count($fields); $i++) {
        $fields[$i]['newid'] = $i;
    }
    //var_dump($fields);
    //wp_die();

    $query = 'SET SESSION group_concat_max_len=45500000';
    $wpdb->get_results($query);

    $t_1_name = $wpdb->prefix . "formmaker_submits";
    $sql = "SELECT group_id as id , date as submit_date,  GROUP_CONCAT( CONCAT(element_label,':',element_value) SEPARATOR  '**$**' ) AS fields " . 
            "FROM  $t_1_name WHERE form_id=$id GROUP BY group_id " .
            "ORDER BY group_id DESC";
    $result_2 = $wpdb->get_results($sql);

    /*var_dump( $result_2);
    wp_die();*/
    //$m = count($result_2) - 1;
    for ($i=0; $i < count($result_2); $i++) {   
        $result_2[$i]->fields = explode( '**$**', $result_2[$i]->fields);
        for ($j=0; $j < count($result_2[$i]->fields); $j++) {
            $arr = explode( ':', $result_2[$i]->fields[$j], 2);
            $key = $arr[0];
            $value = $arr[1];
            $value = str_replace(array('***br***', '<p>', '</p>'), '', $value);
            /*----------------------------------------------------------------*/
            if (!empty($value)) {
                for ( $k = 0; $k < count($fields); $k++) {
                    $str_type = substr($fields[$k]['type'],5);
                    if ($str_type == 'checkbox') $name = 'wdform_'.$key.'_element'.$id*10;
                    else $name = 'wdform_'.$key.'_element'.$id;
                    if ($str_type == 'text') $str_type = 'input';
                    if ($fields[$k]['id'] == $key) {
                        $result_2[$i]->fields[$j] = $fields[$k]['newid'];
                        $nn = $fields[$k]['newid'];
                        $fl_arr[$nn] = array(
                                            'type' => $str_type,
                                            'label'=> $fields[$k]['w_field_label'],
                                            'attributes' => $fields[$k]['w_attr_name'],
                                            'name' => $name,
                                            'value' => $value 
                                        );
                    }
                }
            }
            else $result_2[$i]->fields[$j] = 0;
            /*----------------------------------------------------------------*/
        }
        sort($result_2[$i]->fields, SORT_NUMERIC);
        foreach ($result_2[$i]->fields as $key => $value) {
            foreach ($fl_arr as $k => $val) {
                if ($value == $k) {
                    $result_2[$i]->fields[$key] = $val;
                }
            }
            //remove duplicate values from a multi-dimensional array
            $result_2[$i]->fields = array_map("unserialize", array_unique(array_map("serialize", $result_2[$i]->fields)));
        }
        //$m--;
    }
    $array = array(
        "id" => $id,
        "title" => $result[0] -> title,
        "export date" => date("Y-m-d H:i:s"),
        "items" => $result_2
    );
    return $array;
}
add_action ('admin_init', 'submission_export_to_json_file');

function submission_export_The_Json($arr) { 
}