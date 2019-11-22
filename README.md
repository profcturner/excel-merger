# excel-merger
A very basic tool for merging many source Excel files into a master file, using a simple mapping config file

## Requirements

This script is written for Python 3.6+ and uses OpenPyXL

## Usage

The script can run either interactively (use no command line parameters) or with defined command line parameters. Run

```excel-merger --help```

for more details. It will run through a directory of Excel files, using a mapping file to merge them into one file.

You will need to write your own mapping.cfg - which is a plain text file, to define which cells you want to go where. A sample (and real) mapping.cfg is provided.

## What can it do?

It can set cells in the output spreadsheet to text of your choice, and it can copy cells from any sheet of a source spreadsheet into cells of your choice, appending new rows in the output as required.

There is very basic support for placing mapping instructions in blocks so that the block can be omitted on a condition from the source spreadsheet. For instance, you can use this to skip mapping of regions of an input sheet that are empty.

## Limitations

This is currently very basic, and has no error checking for errors such as missing files.

Likewise, the mapping language is currently very simple.
