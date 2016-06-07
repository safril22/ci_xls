 <?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

class Create_xls extends CI_Controller {

	public function index()
	{
		//load library PHPExcel
		$this->load->library('phpexcel');
		$this->load->library('PHPExcel/IOFactory');

		// merubah style border pada cell yang aktif (cell yang terisi)
		$styleArray = array( 'borders' => 
			array( 'allborders' => 
				array( 'style' => PHPExcel_Style_Border::BORDER_THIN, 'color' => array('argb' => '00000000'), 
					), 
				), 
			);

		// melakukan pengaturan pada header kolom
		$fontHeader = array( 
			'font' => array(
				'bold' => true
			),
			'alignment' => array(
				'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
             	'vertical'   => PHPExcel_Style_Alignment::VERTICAL_CENTER,
             	'rotation'   => 0,
			),
			'fill' => array(
            	'type' => PHPExcel_Style_Fill::FILL_SOLID,
            	'color' => array('rgb' => '6CCECB')
        	)
		);

		//membuat object baru bernama $objPHPExcel
		$objPHPExcel = new PHPExcel();
		$objPHPExcel->getProperties()->setTitle("title")->setDescription("description");

		// data dibuat pada sheet pertama
		$objPHPExcel->setActiveSheetIndex(0); 

		//set header kolom
		$objPHPExcel->getActiveSheet()->setCellValue('B2', 'No.'); 
		$objPHPExcel->getActiveSheet()->setCellValue('C2', 'Nama Lengkap'); 
		$objPHPExcel->getActiveSheet()->setCellValue('D2', 'Alamat');

		// pendefinisian data
		$isi = array(
			array('B' => '1', 'C' => 'Budi Santoso', 'D' => 'Depok'),
			array('B' => '2', 'C' => 'Susi Liana', 'D' => 'Jakarta'),
			array('B' => '3', 'C' => 'Ari Agung', 'D' => 'Jakarta'),
			array('B' => '4', 'C' => 'Ira Mandala', 'D' => 'Surabaya'),
			array('B' => '5', 'C' => 'Joko Dolo', 'D' => 'Depok'),
			array('B' => '6', 'C' => 'Hasan Basri', 'D' => 'Bandung'),
		);
		
		// melakukan pengisian data
		foreach($isi as $k => $v)
		{
			$col = $k + 3;
			foreach($v as $k1 => $v1)
			{
				$column = $k1.$col;
				$objPHPExcel->getActiveSheet()->setCellValue($column, $v1); 
			}
		}

		$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
		$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
		$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);

		$objWorksheet = $objPHPExcel->getActiveSheet();
		$objWorksheet->getStyle('B2:D2')->applyFromArray($fontHeader);
		$objWorksheet->getStyle('B2:'.$column)->applyFromArray($styleArray);

		$objWriter = IOFactory::createWriter($objPHPExcel, 'Excel5');
		$objWriter->save("test_".date('Y-m-d H-i-s').".xls");
	}
}