<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2018/10/25
 * Time: 14:09
 */

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class exportExcel {
    public static $spreadSheet;
    public static $sheet;
    public static $sheetSize=10000;

    public static function export($data, $head) {
        self::setHeader($head);
        set_time_limit(0);
        ini_set ('memory_limit', '800M');
        $sheetIndex = 0;

        /*--------------开始从数据库提取信息插入Excel表中------------------*/
        $item = $data;
        $keys = array_keys($head);
        for ($j=0;$j<100000;$j++) {             //循环设置单元格：
            //$key+2,因为第一行是表头，所以写到表格时   从第二行开始写
            $key = $j%5;
            $tempIndex = intval($j/self::$sheetSize);
            if($tempIndex != $sheetIndex) {
                $sheetIndex = $tempIndex;
                self::$spreadSheet->createSheet($sheetIndex);
                self::$sheet = self::$spreadSheet->setActiveSheetIndex($sheetIndex);
                self::setSheetTitle('list_'.($sheetIndex+1));
                self::setHeader($head);
                //$sheet = $spreadsheet->getActiveSheet($sheetIndex);
            }
            $row = ($j + 2)-($tempIndex*self::$sheetSize);
            $col = 'A';
            for ($i = 65; $i < count($head) + 65; $i++) {     //数字转字母从65开始：
                $col = strtoupper(chr($i));
                self::$sheet->setCellValue($col.$row, $item[$keys[$i - 65]].'-'.$j);
            }
            self::setCellStyle("A{$row}:{$col}{$row}", array(
                'font'=>array(
                    'name'=>'Times New Roman',
                    'color'=>'00000000',//FFFF0000
                    'size'=>'11'
                )
            ));
            self::alignCenter("A{$row}:{$col}{$row}");
        }


        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="temp.xlsx"');
        header('Cache-Control: max-age=0');
        $writer = new Xlsx(self::$spreadSheet);
        $writer->save('php://output');

        //删除清空：
        self::$spreadSheet->disconnectWorksheets();
        //unset(self::$spreadSheet);
        exit;
    }

    public static function init($pro=null){
        self::$spreadSheet = new Spreadsheet();
        self::$sheet = self::$spreadSheet->getActiveSheet();

        self::$spreadSheet->getProperties()
            ->setCreator(isset($pro['author'])?:'author')    //作者
            ->setLastModifiedBy(isset($pro['author'])?:'author') //最后修改者
            ->setTitle(isset($pro['title'])?:'标题')  //标题
            ->setSubject(isset($pro['subject'])?:'') //副标题
            ->setDescription(isset($pro['desc'])?:'')  //描述
            ->setKeywords(isset($pro['keywords'])?:'') //关键字
            ->setCategory(isset($pro['category'])?:''); //分类
    }

    public static function setHeader($head, $attr=null) {
        $head = array_values($head);
        $count = count($head);  //计算表头数量

        $col = '';
        for ($i = 65; $i < $count + 65; $i++) {     //数字转字母从65开始，循环设置表头：
            $col = strtoupper(chr($i));
            self::setColWith($col, isset($attr['with'])?:'30');
            self::$sheet->setCellValue($col . '1', $head[$i - 65]);
            self::alignCenter($col . '1');
        }

        self::setCellStyle('A1:'.$col.'1', array(
            'font'=>array(
                'name'=>isset($attr['font_name'])?:'宋体',
                'size'=>isset($attr['font_size'])?:'14',
                'color'=>isset($attr['font_color'])?:'00000000',
                'bold'=>true
            )
        ));
    }

    public static function setHeaderStyle($col) {

    }

    public static function setSheetTitle($title) {
        self::$sheet->setTitle($title);
    }

    public static function setColWith($col, $width) {
        if($col!='') {
            $sheet = self::$sheet->getColumnDimension($col);
        } else {
            $sheet = self::$sheet->getDefaultColumnDimension();
        }
        if($width=='auto') {
            $sheet->setAutoSize(true);
        } else {
            $sheet->setWidth($width);
        }
    }


    public function setRowHeight($row, $height) {
        if($row) {//设置第10行行高为100pt。
            self::$sheet->getRowDimension($row)->setRowHeight($height);
        } else {
            //设置默认行高。
            self::$sheet->getDefaultRowDimension()->setRowHeight($height);
        }
    }

    /*
     * $style = array(
     *     'font'=>array(
     *          'name'=>'',
     *          'size'=>'',
     *          'color'=>'',
     *          'bold'=>''
     *      )
     * )
     * */
    public static function setCellStyle($cell, $style) {
        //$objs = array_keys($style);
        foreach($style as $pro=>$attrs) {
            switch ($pro) {
                case 'font':
                    $font = self::$sheet->getStyle($cell)->getFont();
                    foreach($attrs as $attr=>$val) {
                        switch($attr) {
                            case 'name':
                                $font->setName($val);
                                break;
                            case 'size':
                                $font->setSize($val);
                                break;
                            case 'color':
                                $font->getColor()->setARGB($val);
                                break;
                            case 'bold':
                                $font->setBold(true);
                                break;
                        }
                    }
                    break;
            }
        }
        //self::$sheet->getStyle('A7:B7')->getFont()->setBold(true)->setName('Arial')->setSize(10);;
    }

    public static function setLink($cell, $link) {
        self::$sheet->getCell($cell)->getHyperlink()->setUrl($link);
    }


    public static function setBorder($cell, $attrs=['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN, 'color' => ['argb' => 'EEEEEEEE']]) {
        $styleArray = [
            'borders' => [
                'outline' => $attrs
            ],
        ];
        self::$sheet->getStyle($cell)->applyFromArray($styleArray);
    }

    public static function mergeCell($range) {
        self::$sheet->mergeCells($range);
    }

    public function alignCenter($cell) {
        $styleArray = [
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical'=>\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
            ],
        ];
        self::$sheet->getStyle($cell)->applyFromArray($styleArray);
    }

}

$head = ['order_sn'=>'订单编号', 'num'=>'商品总数', 'consignee'=>'收货人', 'phone'=>'联系电话', 'detail'=>'收货地址'];
//数据中对应的字段，用于读取相应数据：

exportExcel::init();//run($data, $head, $keys);
exportExcel::export($head, $head);


