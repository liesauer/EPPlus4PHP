<?php

use nulastudio\Document\EPPlus4PHP\ExcelPackage;
use nulastudio\Document\EPPlus4PHP\ExcelConvert;
use nulastudio\Document\EPPlus4PHP\Range;
use nulastudio\Document\EPPlus4PHP\Style\Color;
use nulastudio\Document\EPPlus4PHP\Style\UnderLineType;
use nulastudio\Document\EPPlus4PHP\Style\VerticalAlignmentFont;

$package = new ExcelPackage(__DIR__ . '/test.xlsx');

$worksheet = $package->workBook->workSheets['sheet2'];
$singleCell = $worksheet->cells['A1'];
// $worksheet->addRow([1,2,3,4,5]);
// $package->save();
// die;


// $singleCell->value = new stdClass();

// var_dump($singleCell->value);

// var_dump($worksheet->address);
// var_dump($worksheet->fullAddress);
// var_dump($worksheet->fullAddressAbsolute);

// $singleCell->formula = function() {
//     return microtime(true);
// };

// functions

// enumerable
// class Enumerable/*  implements \Iterator */
// {
//     private $position = 0;
//     private $array = array(
//         "firstelement",
//         "secondelement",
//         "lastelement",
//     );  

//     public function __construct() {
//         $this->position = 0;
//     }

//     public function rewind() {
//         $this->position = 0;
//     }

//     public function current() {
//         return $this->array[$this->position];
//     }

//     public function key() {
//         return $this->position;
//     }

//     public function next() {
//         ++$this->position;
//     }

//     public function valid() {
//         return isset($this->array[$this->position]);
//     }
// }

// $package->addOrReplaceFunction('test', function(array $args, array $context) {
//     // var_dump($args, $context);
//     return $args;
// });
// $package->addOrReplaceFunction('test2', function(array $args, array $context) {
//     // var_dump($args, $context);
//     return 'Y';
// });

// $worksheet->cells['F7:F9']->formula = 'test2(test(1,1.1,"1",false,,A1:A1,A2:A3))';

// var_dump($worksheet->cells['F4']->value);

// $worksheet->cells['F4']->formula = '';

// $worksheet->cells['F4']->value = new \Enumerable();

// var_dump($worksheet->cells['A1']->value);

// $worksheet->cells['A1']->formula = '';

// var_dump($worksheet->cells['A1']->value);

// $package->workBook->workSheets->add('a');
// $package->workBook->workSheets->add('b');
// $package->workBook->workSheets->add('c');

// var_dump(count($package->workBook->workSheets));
// foreach ($package->workBook->workSheets as $name => $worksheet) {
//     var_dump($name, $worksheet);
// }

// exit;


// worksheet
// var_dump(isset($package->workBook->workSheets[1]));
// var_dump(isset($package->workBook->workSheets[2]));
// var_dump(isset($package->workBook->workSheets[3]));
// var_dump(isset($package->workBook->workSheets['test sheet']));
// var_dump(isset($package->workBook->workSheets['test sheet2']));
// var_dump($package->workBook->workSheets['test sheet2']->cells['A1']);

// $package->workBook->workSheets->add('test sheet');

// addressing
// var_dump(ExcelConvert::toName(100));
// var_dump(ExcelConvert::toName(9999));
// var_dump(ExcelConvert::toName(1048576));
// var_dump(ExcelConvert::toIndex('A'));
// var_dump(ExcelConvert::toIndex('XFD'));
// $tmp_addr;
// var_dump(Range::tryParseAddress("a", $tmp_addr), $tmp_addr);
// var_dump(Range::tryParseAddress("a", $tmp_addr), $tmp_addr); // y
// var_dump(Range::tryParseAddress("xfd", $tmp_addr), $tmp_addr); // y
// var_dump(Range::tryParseAddress("xfe", $tmp_addr), $tmp_addr); // n
// var_dump(Range::tryParseAddress("0", $tmp_addr), $tmp_addr);   // n
// var_dump(Range::tryParseAddress("1", $tmp_addr), $tmp_addr);   // y
// var_dump(Range::tryParseAddress("88888888", $tmp_addr), $tmp_addr);   // n
// var_dump(Range::tryParseAddress("a:a", $tmp_addr), $tmp_addr); // y
// var_dump(Range::tryParseAddress("1:1", $tmp_addr), $tmp_addr); // y
// var_dump(Range::tryParseAddress("da", $tmp_addr), $tmp_addr);  // y
// var_dump(Range::tryParseAddress("a,1,1:1,a2,a2:a5", $tmp_addr), $tmp_addr); // y
// var_dump(Range::tryParseAddress("a,1,1:1,a2,88888888,1q:q1,a2:a5", $tmp_addr), $tmp_addr); // n

// style font color
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->color = '#AABBCC';
// $color = $package->workBook->workSheets['test sheet']->cells['A1']->style->font->color;
// var_dump($color);
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->size = 50;

// $color->alpha = 0xBB;
// exit;
// style font underLineType
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLine = true;
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLineType = UnderLineType::Double;
// var_dump($package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLineType);
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLine = false;
// var_dump($package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLineType);
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLineType = UnderLineType::SingleAccounting;
// var_dump($package->workBook->workSheets['test sheet']->cells['A1']->style->font->underLineType);

// style font verticalAlignmentFont
// bug: effective only for the first letter
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->verticalAlign = VerticalAlignmentFont::Superscript;
// var_dump($package->workBook->workSheets['test sheet']->cells['A1']->style->font->verticalAlign);

// style fill
// $package->workBook->workSheets['test sheet']->cells['A1']->style->fill->backgroundColor = 0x00BBCCDD;
// var_dump($package->workBook->workSheets['test sheet']->cells['A1']->style->fill->backgroundColor);

// style border
// use nulastudio\Document\EPPlus4PHP\Style\BorderStyle;
// $singleCell->style->border->top->style = BorderStyle::Dotted;
// $singleCell->style->border->top->color = 0x00BBCCDD;
// $singleCell->style->border->bottom->style = BorderStyle::Dotted;
// $singleCell->style->border->bottom->color = Color::GREEN_COLOR;
// $singleCell->style->border->diagonal->style = BorderStyle::Dashed;
// $singleCell->style->border->diagonal->color = Color::BLUE_COLOR;
// $singleCell->style->border->diagonalUp = true;
// $singleCell->style->border->diagonalDown = true;

// style numberformat
// $singleCell->style->numberFormat->format = 'yyyy/m/d h:mm';

// reading and writting
// single cell
// var_dump($worksheet->cells['A1']->value);
// multi rows
// var_dump($worksheet->cells['1']->value);
// foreach ($worksheet->datas->value as $value) {
//     var_dump($value);
// }
// multi columns
// var_dump($worksheet->datas->value);
// var_dump($worksheet->cells['A1:A5']->value);
// $worksheet->cells['1:3']->value = [
//     [1,2,3,4,5,6,7,8,9],
//     [9,9,9],
//     [0,0,0,0,0,9],
// ];
// $package->save();
// var_dump(count($data));
// var_dump(count($data[0]));
// // mutli datas
// var_dump($worksheet->cells['A1:C20']->value);

// merge
// $worksheet->cells['A1:C1']->merge = true;
// $worksheet->cells['A1:C2']->merge = true;

// alignment
// use nulastudio\Document\EPPlus4PHP\Style\HorizontalAlignment;
// use nulastudio\Document\EPPlus4PHP\Style\VerticalAlignment;
// $worksheet->cells['A1:C2']->style->horizontalAlignment = HorizontalAlignment::Right;
// $worksheet->cells['A1:C2']->style->verticalAlignment = VerticalAlignment::Top;


// $package->save("test2");

// save/saveAs
// $package->saveAs(__DIR__ . '/test2.xlsx');
// $package->saveAs(__DIR__ . '/test2_pwd.xlsx', 'test');


// add row
// $datas = [
//     [1,2,3,4,5,6,7],
//     [8,8,8,8,8,8,8,8],
//     [2,2,2,2,2,2,2,2],
//     ['1','','qq','wa','asd','ppp'],
// ];
// foreach ($datas as $data) {
//     $worksheet->addRow($data);
// }

// ExcelPackage::foo(new DateTime());
// insert
// $worksheet->insertRow(3, [6,6,6,6,6,6]);
// $worksheet->insertColumn(5, [9,9,9,9,9,9]);

// add column
// $datas = [
//     ['|',1, 2, 3, 4, 5, 6, 7],
//     ['|',8, 8, 8, 8, 8, 8, 8, 8],
//     ['|',2, 2, 2, 2, 2, 2, 2, 2],
//     ['|','1', '', 'qq', 'wa', 'asd', 'ppp'],
// ];
// foreach ($datas as $data) {
//     $worksheet->addColumn($data);
// }

// var_dump($package);

// comment
// $singleCell->comment = "你好";
// $singleCell->comment->text = null;
// $singleCell->comment->author = "LiesAuer";
// $singleCell->comment = null;
// $singleCell->comment = null;
// $singleCell->comment = 0;

// ArrayAccess
// var_dump($package->workBook->workSheets[1]);
// var_dump($package->workBook->workSheets[2]);
// var_dump($package->workBook->workSheets['test sheet']);
// var_dump($package->workBook->workSheets['test sheet2']);

// Countable
// var_dump(count($package->workBook->workSheets));

// IEnumerable
// foreach ($package->workBook->workSheets as $name => $worksheet) {
//     var_dump($name/* , $worksheet */);
// }
// foreach ($package->workBook->workSheets as $name => $worksheet) {
//     var_dump($name/* , $worksheet */);
// }

// addressing and cells

// $range = $package->workBook->workSheets['test sheet']->cells['A1']->style->font->size = 50;

// var_dump($range);

// $package->workBook->workSheets['test sheet']->cells['A1']->font->size = 50;

// ??????

$worksheet->addRow(
    false,
    666
    // true,
    // 666.0,
    // 'haha',
    // ['a','b','c'],
    // new stdClass,
    // null,
    // function(){},
    // curl_init()
);

echo '';

$package->save();