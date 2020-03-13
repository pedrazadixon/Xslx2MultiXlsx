<?php

require 'vendor/autoload.php';

class Xslx2MultiXlsx
{

    public $inptu_file_path = NULL;
    public $has_titles = NULL;
    public $traslator_file_path = NULL;
    public $template_path = NULL;

    function __construct($inptu_file_path, $traslator_file_path, $template_path, $has_titles = NULL)
    {
        $this->inptu_file_path = $inptu_file_path;
        $this->traslator_file_path = $traslator_file_path;
        $this->template_path = $template_path;
        $this->has_titles = ($has_titles === NULL) ? false : true;
    }

    function loadTemplate($file)
    {
        if (!file_exists($file))
            exit($file . ' not exists');

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $template_spreadsheet = $reader->load($this->template_path);

        return $template_spreadsheet;
    }


    function writeOutputFiles()
    {
        $spreadsheet = $this->loadTemplate($this->template_path);
        $sheet = $spreadsheet->getActiveSheet();
        $input_data = $this->getInputData();
        $positions = $this->getPositions();

        ##### comprobar titulos flag
        unset($input_data[0]);

        foreach ($input_data as $row_n => $row) {

            foreach ($positions as $col_title => $data) {

                if (is_array($data["target_cells"])) {
                    foreach ($data["target_cells"] as $pos) {
                        $sheet->setCellValue($pos, $row[$data["input_data_col_index"]]);
                    }
                    continue;
                }

                $sheet->setCellValue($data["target_cells"], $row[$data["input_data_col_index"]]);
            }

            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
            $writer->save('output/file_' . $row_n . '.xlsx');
            print_r('Created file: output/file_' . $row_n . '.xlsx' . PHP_EOL);
        }
    }

    function getPositions()
    {
        $file_titles = $this->getTitles();
        $traslator = $this->getTraslator();
        $titles_positions = [];
        foreach ($traslator as $traslator_title => $cells) {
            $titles_positions[$traslator_title] = [
                'input_data_col_index' => array_search($traslator_title, $file_titles),
                'target_cells' => $cells,
            ];
        }
        return $titles_positions;
    }


    function getInputData()
    {
        if (!file_exists($this->inptu_file_path))
            exit($this->inptu_file_path . ' not exists');

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);
        $input_spreadsheet = $reader->load($this->inptu_file_path);
        $input_worksheet = $input_spreadsheet->getActiveSheet();
        $raw_data = $input_worksheet->toArray();
        return $raw_data;
    }

    function checkTitles()
    {
        $file_titles = $this->getTitles();
        $traslator = $this->getTraslator();

        $not_found_cols = '';

        foreach ($traslator as $traslator_title => $cells) {

            if (!in_array($traslator_title, $file_titles)) {
                $not_found_cols .= $traslator_title . '|';
            }
        }

        if ($not_found_cols == '')
            return true;

        $not_found_cols = '|' . $not_found_cols . ' column(s) in ' . $this->traslator_file_path . ' not found in ' . $this->inptu_file_path;
        exit($not_found_cols);
    }

    function getTitles()
    {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(true);
        $input_spreadsheet = $reader->load($this->inptu_file_path);
        $input_worksheet = $input_spreadsheet->getActiveSheet();

        $titles = [];
        foreach ($input_worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            foreach ($cellIterator as $cell) {
                $titles[] = $cell->getValue();
            }
            break;
        }

        return $titles;
    }

    function getTraslator()
    {
        $raw_json = @file_get_contents($this->traslator_file_path);
        if ($raw_json === false)
            exit($this->traslator_file_path . ' not found or permisions denied');

        $json = json_decode($raw_json, true);
        if ($json === NULL)
            exit($this->traslator_file_path . ' incorrect format or very large');

        return $json;
    }
}

function dd($dato)
{
    var_dump($dato);
    exit();
}
