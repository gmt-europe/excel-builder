# Excel Builder

LGPL License.

## Introduction

Excel Builder is an application that allows you to create simple Excel files from JSON
input. The purpose of this application is to allow building of Excel files from platforms
that do not support integration with e.g. Apache (N)POI. Instead, such a platform can
use this application to convert a simple JSON format into an Excel file.

See the [wiki](https://github.com/gmt-europe/excel-builder/wiki) for documentation on the JSON format
Excel Builder expects as input.

## Command line arguments

Excel Builder takes the following command line arguments:

* `-s <sheet name>`: The name of the workbook sheet. If this parameter is omitted, "Sheet1" is used instead;
* `-f <font>`: The default font. If this parameter is omitted, "Calibri" is used instead;
* `-S <font size>`: The default font size. If this parameter is omitted, 11pt is instead;
* `-i <input file name>`: The file name to use as input. If this parameter is omitted,
  `STDIN` is used instead.
* `-o <output file name>`: The file name to output the Excel file to. If this parameter
  is omitted, `STDOUT` is used instead.
* `-d <date format>`: The date format used to format dates. If this parameter is omitted, "m/d/yyyy h:mm" instead.

## Bugs

Bugs should be reported through github at
[http://github.com/gmt-europe/excel-builder/issues](http://github.com/gmt-europe/excel-builder/issues).

## License

Excel Builder is licensed under the LGPL 3.
