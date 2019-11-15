<?php
namespace App\Http\Controllers;

use PhpOffice\PhpWord\TemplateProcessor;
use Illuminate\Http\Request;

class ReportsController extends Controller{

public function word(){
	$templateProcessor = new TemplateProcessor('./templates/Certificate of Recognition.docx');

	$templateProcessor->setValue('first_name', 'John');
	$templateProcessor->setValue('last_name', 'Dy');

	$templateProcessor->saveAs('John Dy Certificate.docx');

	return response()->download('John Dy Certificate.docx');
}
public function excel(){
	$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('template/.Form138.xlsx');

		$worksheet = $spreadsheet->getActiveSheet();

		$worksheet->getCell('A7')->setValue('Name: John Dy');
		$worksheet->getCell('A7')->setValue('11-B');

		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xls');
		$writer->save('Form138.xls');

		return response()->download('Form138.xls');
		}
	
}


