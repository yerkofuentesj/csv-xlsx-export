<?php
defined('BASEPATH') OR exit('No direct script access allowed');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
class Welcome extends CI_Controller {

	public function __construct()
	{
		parent::__construct();
		$this->db = $this->load->database('NAME_HOST', true);
	}

	/**
	 * Index Page for this controller.
	 *
	 * Maps to the following URL
	 * 		http://example.com/index.php/welcome
	 *	- or -
	 * 		http://example.com/index.php/welcome/index
	 *	- or -
	 * Since this controller is set as the default controller in
	 * config/routes.php, it's displayed at http://example.com/
	 *
	 * So any other public methods not prefixed with an underscore will
	 * map to /index.php/welcome/<method_name>
	 * @see https://codeigniter.com/userguide3/general/urls.html
	 */
	public function index()
	{
		$spreadsheet = new Spreadsheet();
		$sheet = $spreadsheet->getActiveSheet();

		foreach (range('A','E') as $coulumID) 
		{
			$spreadsheet->getActiveSheet()->getColumnDimension($coulumID)->setAutosize(true);
		}
		$sheet->setCellValue('A1','ID');
		$sheet->setCellValue('B1','USERNAME');
		$sheet->setCellValue('C1','NAME');
		$sheet->setCellValue('D1','GENDER');
		$sheet->setCellValue('E1','EMAIL');

		$users = $this->db->query("SELECT * FROM dbo.PEOPLE")->result_array();
		$x=2;
		foreach ($users as $row) 
		{
			$sheet->setCellValue('A'.$x, $row['id']);
			$sheet->setCellValue('B'.$x, $row['username']);
			$sheet->setCellValue('C'.$x, $row['name']);
			$sheet->setCellValue('D'.$x, $row['gender']);
			$sheet->setCellValue('E'.$x, $row['email']);
			$x++;
		}

		$writer = new Xlsx($spreadsheet);
		$fileName='users_details_exports.xlsx'; // ejemplo.xlsx o .csv
		//$writer->save($fileName); // guardar 

		header('Content-Type: appliction/vmd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="'.$fileName.'"');
		$writer->save('php://output');
	}
}
