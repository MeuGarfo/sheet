<?php
namespace Basic;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Common\Type;
class Sheet{
    function alpha_to_num($a){
        $l = strlen($a);
        $n = 0;
        for($i = 0; $i < $l; $i++)
        $n = $n*26 + ord($a[$i]) - 0x60;
        return $n-1;
    }
    function num_to_alpha($n){
        for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
        $r = chr($n%26 + 0x61) . $r;
        return $r;
    }
    function sheet_to_alpha($sheet){
        $fixed_sheet=false;
        foreach($sheet as $key=>$value){
            foreach($value as $value_key=>$value_value){
                unset($value[$value_key]);
                $value[$this->num_to_alpha($value_key)]=$value_value;
            }
            unset($sheet[$key]);
            $fixed_sheet[$key+1]=$value;
        }
        return $fixed_sheet;
    }
    function to_array($sheet_name){
        //le a lista
        $ext=pathinfo($sheet_name,PATHINFO_EXTENSION);
        switch($ext){
            case 'csv':
            $reader = ReaderFactory::create(Type::CSV); // for CSV files
            break;
            case 'ods':
            $reader = ReaderFactory::create(Type::ODS); // for ODS files
            break;
            case 'xlsx':
            $reader = ReaderFactory::create(Type::XLSX); // for XLSX files
            break;
        }
        $reader->open($sheet_name);
        foreach ($reader->getSheetIterator() as $sheet) {
            $sheetName = $sheet->getName();
            foreach ($sheet->getRowIterator() as $row) {
                $list['lists'][$sheetName][]=$row;
            }
        }
        $reader->close();
        $lists=false;
        //adiciona letras nas colunas
        foreach ($list['lists'] as $key => $value) {
            $lists[$key]=$this->sheet_to_alpha($value);
        }
        //retorna a lista
        return $lists;
    }
}
