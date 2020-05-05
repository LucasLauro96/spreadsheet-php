<?php

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

require 'vendor/autoload.php';

$conexao = new conexao();
$con = $conexao->conecta();

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('spreadsheet.xlsx');

$worksheet = $spreadsheet->getActiveSheet();

$worksheet->getCell('A' . '1')->setValue('Your Value');

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('spreadsheet.xlsx');
