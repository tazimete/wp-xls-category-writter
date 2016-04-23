
<?php 
	/*
	Plugin Name: xls-writer
	Plugin URI: http://wordpress.org/plugins/hello-dolly/
	Description: This is not just a plugin, it symbolizes the hope and enthusiasm of an entire generation summed up in two words sung most famously by Louis Armstrong: Hello, Dolly. When activated you will randomly see a lyric from <cite>Hello, Dolly</cite> in the upper right of your admin screen on every page.
	Author: NeshmediaBD
	Version: 1.0
	Author URI: http://ma.tt/
	*/
	
	if(!class_exists('PHPExcel')){
		require_once('phpExcel/PHPExcel.php');
	}
	
	ini_set('max_execution_time', -1);
	ini_set('memory_limit', -1);
	
	//Registering Arabic Tutor menu in Admin Page 
	add_action( 'admin_menu', 'register_xlsw_menu' );

	function register_xlsw_menu() {
		add_menu_page( 'xls-writer', ' XLS Writer', 'manage_options', 'xls_writer','show_xlsw_main_page' ); 	
	}
	
	function show_xlsw_main_page(){ ?>
		<div class="wrap">
			<div class="xls-witer-container">
				<h2>XLS Writer</h2>
				<div class="xsl-writer-form">
					<form action="#" method="post" id="xls-witer-form" enctype="multipart/form-data">
						<table>
							<tr><td><label for="xls-uploader"> Upload XLS File </label></td></tr>
							<tr><td> <input type="file" name="xls-uploader"/></td></tr>
							<tr><td><input type="submit" name="xls-uploader-submit" value="Upload" class="button button-primary button-large"/></td></tr>
						</table>
					</form>
				</div>
			</div>
		</div>
	<?php }
	
	function receive_xls_file(){
		if(isset($_FILES) && isset($_POST['xls-uploader-submit'])){
				$inputFileName = $_FILES['xls-uploader']['tmp_name'];
				//$inputFileType = $_FILES['xls-uploader']['type'];
				/* $objPHPExcelForWriter = PHPExcel_IOFactory::createReader('Excel2007');
				$objPHPExcelForWriter = $objPHPExcelForWriter->load($inputFileName);
				$objPHPExcelForWriter->setActiveSheetIndex(0);
				$row = 1;
				$objPHPExcelForWriter->getActiveSheet()->SetCellValue('K1', "Category");
				 */
				//print_r($_FILES); exit;
				
				if(is_uploaded_file($_FILES['xls-uploader']['tmp_name'])){
					//echo "File Uploaded Successfully fdbgfdgbfdgb bfd gdf gdfg";
				}else{	
					//print "Failed to upload file.................................................";
				}
					
				  try {
					$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
					$objReader = PHPExcel_IOFactory::createReader($inputFileType);
					$objPHPExcel = $objReader->load($inputFileName);
					$objPHPExcel->setActiveSheetIndex(0);
					$objPHPExcel->getActiveSheet()->SetCellValue('K1', "Category");
				} catch(Exception $e) {
					die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
				} 
				
				//  Get worksheet dimensions
				$sheet = $objPHPExcel->getSheet(0);
				$highestRow = $sheet->getHighestRow();
				$highestColumn = $sheet->getHighestColumn();
				$file = fopen("category_list.csv","w");

				///$category_list= array();
				//$post_id= array();
				
				//  Loop through each row of the worksheet in turn
				for ($row = 2; $row <= $highestRow; $row++) {
					//  Read a row of data into an array
					$rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, 
						NULL, TRUE, FALSE);
					$product_code = trim($rowData[0][1]);
			
					 
					//$product = new WP_Query( $rd_args ); 
					
						global $wpdb;
						/* $meta = $wpdb->get_results("SELECT * FROM `".$wpdb->postmeta."` WHERE meta_key='mirror-product-code' AND meta_value='".$product_code."'");
						if (is_array($meta) && !empty($meta) && isset($meta[0])) {
							$meta_id = $meta[0];
						}	
						if (is_object($meta_id)) {
							$post_id[] =  $meta_id->post_id;
						} */
						
						 
						$args = array(
							'meta_key'         => 'mirror-product-code',
							'meta_value'       => $product_code,
							'post_type'        => 'mirror',
						);
						$posts_array = get_posts( $args );
						wp_reset_query();
						
						if(!$posts_array){
							$args = array(
							'meta_key'         => 'mouldings-product-code',
							'meta_value'       => $product_code,
							'post_type'        => 'frame',
							);
							$posts_array = get_posts( $args );
							wp_reset_query();
						}
					
					if ( $posts_array ) {
						$post = $posts_array[0];
						$post_id = $post->ID;
					}
						
						//print_r($meta); exit;
					
					//print_r($meta); exit;
					//print_r($post_id); exit;
					
					if($post_id){
						//$category_list = get_the_category_list( ",", "multiple", $post_id );
						//$category_list[] = get_the_term_list( $post_id, 'mirror_cat', " ", ",", " " );
						$category_list = get_the_terms( $post_id, 'mirror_cat' );
						//$category_list[] = get_the_category( $post_id );
						//echo '<pre>'; print_r($product); exit; echo '</pre>';
						//print_r($category_list); exit;
						if(!$category_list){
							$category_list = get_the_terms( $post_id, 'mouldings_cat');
							//$category_list[] = get_the_category( $post_id);
							//$category_list[$row-2] = get_the_term_list( $post_id, 'mouldings_cat', " ", ",", " " );
							//$category_list = get_the_category_list( ",", "multiple", $post_id );
						}
						
					}
					//print_r($category_list); exit;
					$objPHPExcel->getActiveSheet()->SetCellValue('K'.$row, $category_list[0]->name);
					$lines[][0] = $category_list[0]->name;
					foreach ($lines as $line){
						fputcsv($file, $line);
					}

					
				} 
				
				//echo '<pre>'; print_r($category_list); echo '</pre>'; exit;
				//$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcelForWriter);
				//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
				$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel, 'Excel2007');
				$objWriter->save("new_cat_list.xls");	
				fclose($file);
			} else {
				//echo "<h2>Error loading file</h2>";
			}	
	}
	
	add_action('init', 'receive_xls_file');
	
	
	
?>