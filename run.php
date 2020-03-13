<?php

require_once 'Xslx2MultiXlsx.php';

$input_folder = "input";
$template_folder = "templates";

function run()
{
    $file = menu();

    if (findTraslatorAndTemplate($file)) {

        $process = new Xslx2MultiXlsx(
            $GLOBALS["input_folder"] . "/" . $file,
            $GLOBALS["input_folder"] . "/" . pathinfo($file, PATHINFO_FILENAME) . ".json",
            $GLOBALS["template_folder"] . "/" . $file,
            true
        );

        $process->writeOutputFiles();
    }
}

function findTraslatorAndTemplate($file)
{

    $json_file = pathinfo($file, PATHINFO_FILENAME) . ".json";
    if (!file_exists($GLOBALS["input_folder"] . "/" . $json_file))
        exit($json_file . " traslator file not found in " . $GLOBALS["input_folder"] . " folder");

    $template_file = pathinfo($file, PATHINFO_FILENAME) . ".xlsx";
    if (!file_exists($GLOBALS["template_folder"] . "/" . $template_file))
        exit($template_file . " template file not found in " . $GLOBALS["template_folder"] . " folder");

    return true;
}

function printMenu($file_list)
{
    printf("\r\nInput data files found in %s folder:\r\n\r\n", $GLOBALS["input_folder"]);

    for ($i = 0; $i < count($file_list); $i++) {
        printf("\t[%d] %s\r\n", ($i + 1), $file_list[$i]);
    }
    printf("\r\n");

    $selected = readline("Enter number of a file: ");
    return $selected;
}

function menu()
{
    $file_list = getFileList($GLOBALS["input_folder"]);

    $invalid_menu = true;
    do {
        $option = printMenu($file_list);

        if (@isset($file_list[($option - 1)])) {
            $invalid_menu = false;
            $selected_file = $file_list[($option - 1)];
        } else {
            printf("\r\ninvalid option!\r\n");
        }
    } while ($invalid_menu);

    return $selected_file;
}

function getFileList()
{
    $all_files = @scandir($GLOBALS["input_folder"]);

    if ($all_files === false)
        exit($GLOBALS["input_folder"] . ' folder not exists');

    $xlsx_files = array_filter($all_files, function ($file) {
        if (pathinfo($file, PATHINFO_EXTENSION) == "xlsx")
            return $file;
    });

    return array_values($xlsx_files);
}

run();
