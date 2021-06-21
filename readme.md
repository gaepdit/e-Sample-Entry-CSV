# XL-ESE – eSample Data Entry Spreadsheet and Submission File Generator

eSample Entry (e-SE) is a proprietary application used by many state Drinking Water programs to accept the electronic submission of drinking water lab data. Users can enter the data directly into e-SE or upload the data as XML files. This project is an effort to simplify the generation of XML data files for use in e-SE.

We have created a Microsoft Excel spreadsheet with data validation rules applied and a tool to export the data as a valid file for submission. Because the spreadsheet uses VBA (macros) for some functionality, the user must enable content in the security warning displayed in Excel.

## Development

VBA code is exported to separate files using a Git pre-commit hook as described here:
[How to use Git hooks to version-control your Excel VBA code](https://www.xltrail.com/blog/auto-export-vba-commit-hook). To enable, follow these steps:

1. [Install Python](https://www.python.org/).

2. Install [oletools](https://github.com/decalage2/oletools) version 0.53.1 by running the following command:

    `pip install oletools`

3. Enable the Git pre-commit hook by running the following command:

    `git config core.hooksPath .githooks`

Alternatively, the VBA files can be exported manually using the `export-vba.bat` file.
