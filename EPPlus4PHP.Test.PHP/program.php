<?php

use nulastudio\Document\EPPlus4PHP\ExcelPackage;
use nulastudio\Document\EPPlus4PHP\Range;

// $package = new ExcelPackage(__DIR__ . '/test.xlsx');


// worksheet
// var_dump(isset($package->workBook->workSheets[1]));
// var_dump(isset($package->workBook->workSheets[2]));
// var_dump(isset($package->workBook->workSheets[3]));
// var_dump(isset($package->workBook->workSheets['test sheet']));
// var_dump(isset($package->workBook->workSheets['test sheet2']));
// var_dump($package->workBook->workSheets['test sheet2']->cells['A1']);

// $package->workBook->workSheets->add('test sheet');

// addressing
var_dump(Range::parseAddress("a"));
var_dump(Range::parseAddress("1"));
var_dump(Range::parseAddress("a:a"));
var_dump(Range::parseAddress("da"));
var_dump(Range::parseAddress("1:1"));
var_dump(Range::parseAddress("a,1,1:1,a2,a2:a5"));

// $package->save();

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