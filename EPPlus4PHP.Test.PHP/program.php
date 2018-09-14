<?php

use nulastudio\Document\EPPlus4PHP\ExcelPackage;
use nulastudio\Document\EPPlus4PHP\ExcelConvert;
use nulastudio\Document\EPPlus4PHP\Range;
use nulastudio\Document\EPPlus4PHP\Style\Color;
use nulastudio\Document\EPPlus4PHP\Style\UnderLineType;
use nulastudio\Document\EPPlus4PHP\Style\VerticalAlignmentFont;

$package = new ExcelPackage(__DIR__ . '/test.xlsx');

$worksheet = $package->workBook->workSheets['test sheet'];
$singleCell = $worksheet->cells['A1'];

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
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->color = 0x00BBCCDD;
// $color = $package->workBook->workSheets['test sheet']->cells['A1']->style->font->color;
// // var_dump($color);
// $package->workBook->workSheets['test sheet']->cells['A1']->style->font->size = 50;

// $color->alpha = 0xBB;

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
$worksheet->cells['1:3']->value = [
    [1,2,3,4,5,6,7,8,9],
    [9,9,9],
    [0,0,0,0,0,9],
];
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


$package->save();

// var_dump($package);

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


// $package->save();