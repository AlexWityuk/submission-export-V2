<?php
if( ! class_exists( 'WP_List_Table' ) ) {
    require_once( ABSPATH . 'wp-admin/includes/class-wp-list-table.php' );
}
class Export_Forms_List_Table extends WP_List_Table {
	function get_columns() {
	  	$columns = array(
	  		'cb'        => '<input type="checkbox" />',
		    'title' => 'Title',
		    'id'    => 'Shortcode',
		    'form_fields' => 'Select the category'
	  	);
	  	return $columns;
	}
    function column_cb($item){
        return sprintf(
            '<input type="checkbox" name="exptojson[]" value="%2$s" />',
            /*$1%s*/ $item->title, 
            /*$2%s*/ $item->id
        );
    }
    function column_id($item){
        return sprintf(
            '[Form id="%1$s"]',
            /*$1%s*/ $item->id
        );
    }
    /*function get_bulk_actions() {
	  	$actions = array(
    		'export'    => 'To export the JSON'
	  );
	  return $actions;
	}*/
	function column_default($item, $column_name)
	{
		switch ($column_name) {
			case 'title':
			case 'id':
				return $item->$column_name;
			default:
				return "col name = $column_name ," . print_r($item, true);
		}
	}
    function column_title($item){    
        $actions = array(
            'edit' => sprintf('<a href="?page=%s&action=%s&form=%s">To export the JSON</a>',$_REQUEST['page'],'edit',$item->id)
        );
        return sprintf('%1$s <span style="color:silver">%2$s',
            /*$1%s*/ $item->title,
            /*$2%s*/ $this->row_actions($actions)
        );
    }
    function column_form_fields($item){
        $category_array = $this -> submission_export_Get_categories_name_array ($item->id);
        //var_dump($category_array);
        //wp_die();
        /*
        $str = '<select class="form-control" name="'.$item->id.'expformoption">' .
                '<option disabled selected value>Select a category</option>';
                foreach ($category_array as $key => $value) {
                	$str = $str . "<option>$value</option>";
                }
                       
        $str = $str . '</select>';
        */
        /*-------------------------------------------------------------------*/
        /* select multiple categories in the top legal form submission export 
        /* and compine this categroies into one single Word file
        /*-------------------------------------------------------------------*/
        ?>
    	<style type="text/css">
		.column-form_fields div.tabs-panel {
		    min-height: 42px;
		    max-height: 200px;
		    overflow: auto;
		    padding: 0 .9em;
		    border: 1px solid #ddd;
		    background-color: #fdfdfd;
		}
		</style>
        <?php
        $str = '<div class="tabs-panel">
		<ul class="categorychecklist form-no-clear">';
        foreach ($category_array as $key => $value) {
        	$str = $str . '<li id="" class="wpseo-term-unchecked">
				<label class="selectit">
				<input value="'.$value.'" type="checkbox" name="'.$item->id.'expformoption[]" id=""/>'.
					$value.
				'</label>
			</li>';
        }
      	$str = $str . '</ul></div>';             
        return $str;
    }
	function prepare_items() {
		global $wpdb; 
        $table_name = $wpdb->prefix . "formmaker";
        $where_search = "";
		if (isset($_REQUEST['s']) && !empty($_REQUEST['s'])) {
			$search = $_REQUEST['s'];
			$where_search = "WHERE date LIKE '%{$search}%'
			OR text LIKE '%{$search}%'
			OR theme LIKE '%{$search}%'
			";
		}
        $query = "SELECT * FROM $table_name {$where_search}";
     	$data = $wpdb->get_results($query);

	  	$columns = $this->get_columns();
	  	$hidden = array();
	  	$sortable = array();
	  	$this->_column_headers = array($columns, $hidden, $sortable);
	  	$this->items = $data;
  	 	$this->process_bulk_action();
	}
	function process_bulk_action() {
       	$action = $this->current_action();
	    if( 'edit'===$this->current_action() ) {
        	//wp_die('Items exported to JSON!');
        	ob_end_clean();
        	$array = submission_export_Get_Submission($_GET['form']);
        	$file_name = 'tablexport.json';
	        header('Content-Type: application/json; charset=utf-8');
	        echo  json_encode($array, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
	        header("Content-disposition: attachment; filename=".$file_name);
	        exit();
	    }
        return;
    }
    /*----------------------------------------------------------*/
    function submission_export_Get_categories_name_array ($id) {
	    global $wpdb;
	    $table_name = $wpdb->prefix . "formmaker";
	    $query = "SELECT title, form_fields FROM $table_name WHERE id = $id";
	    $result = $wpdb->get_results($query);
	    $arr = explode('*:*new_field*:*', $result[0] -> form_fields );

	    $res_array_cat_name =  array();
	    $fields = array();
	    for ( $i = 0; $i < count($arr); $i++ ) {
	        $arr_val = explode( '*:*', $arr[$i] );
	        $fields[] = $arr_val;
	    }
	    for ( $i = 0; $i < count($fields); $i++) {
	        $item_obj = array();
	        $sub_arr = array();
            $cat_name; //an category name
	        for ( $j = 0; $j < count($fields[$i]); ) {
	            $key = $fields[$i][$j + 1];
	            $value = $fields[$i][$j];
	            if ($key == 'w_field_label') {
	                $cat_name = $value;
	            }
	            if ($key == 'w_attr_name' && $value == 'checkhead=head') {
	                $res_array_cat_name[] = $cat_name;
	            }
	            $j+=2;
	        }
	    }
	    return $res_array_cat_name;
	}
}