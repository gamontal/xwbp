#!/usr/bin/env node

'use strict';
var prog = require('commander');
var pkg = require('./package.json');
var XLSX = require('xlsx');

var workbook;
function readWorkbook(wb) { return XLSX.readFile(workbook); }

function to_json(wb) {
    workbook = wb;
    workbook = readWorkbook(workbook);

    var result = {};
	  workbook.SheetNames.forEach(function(sheetName) {
		    var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		    if(roa.length > 0) {
			      result[sheetName] = roa;
		    }
	  });
    result = JSON.stringify(result, 2, 2);
	  console.log(result);
}

function to_csv(wb) {
    workbook = wb;
    workbook = readWorkbook(workbook);

    var result = [];
	  workbook.SheetNames.forEach(function(sheetName) {
		    var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
		    if(csv.length > 0) {
			      result.push('SHEET: ' + sheetName);
			      result.push('');
			      result.push(csv);
		    }
	  });
	  console.log(result.join('\n'));
}

function to_formulae(wb) {
    workbook = wb;
    workbook = readWorkbook(workbook);

    var result = [];
	  workbook.SheetNames.forEach(function(sheetName) {
		    var formulae = XLSX.utils.get_formulae(workbook.Sheets[sheetName]);
		    if(formulae.length > 0) {
			      result.push('SHEET: ' + sheetName);
			      result.push('');
			      result.push(formulae.join('\n'));
		    }
	  });
	  console.log(result.join('\n'));
}

prog
    .usage('<options> <FILE>')
    .version(pkg.version)
    .option('--json <file>', 'convert to JSON format', to_json)
    .option('--csv <file>', 'convert to CSV format', to_csv)
    .option('--formulae <file>', 'convert to FORMULAE format', to_formulae)

prog.on('--help', function() {
    console.log('  Supported read formats:');
    console.log('');
    console.log('    - Excel 2007+ XML Formats (XLSX/XLSM)');
    console.log('    - Excel 2007+ Binary Format (XLSB)');
    console.log('    - Excel 2003-2004 XML Format (XML "SpreadsheetML")');
    console.log('    - Excel 97-2004 (XLS BIFF8)');
    console.log('    - Excel 5.0/95 (XLS BIFF5)');
    console.log('    - OpenDocument Spreadsheet (ODS)');
    console.log('');
});

prog.parse(process.argv);
