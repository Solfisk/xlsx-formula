# XLSX formulas

Calculation of basic formulas found in Excel sheets.

Note that this module *only* supports basic calculations. The built in functions provided by Excel are not
supported. Also note that this code uses eval. Don't use this on worksheets from untrusted sources.

Example:

Assume my-workbook.xlsx contains

A1    10
A2    20
A3    =A1+A2

And A3 has the name "result" (see cell names in the documentation if you haven't heard about named cells).

    npm install xlsx-formulas
    node -e 'require('xlsx-formulas')('my-workbook.xlsx')
      .then(wb => {
        console.log('Sheet1.A3: ' + wb.Sheet1.A3());
        // outputs 30
        wb.Sheet1.A1 = () => 22;
        console.log('Sheet1.A3: ' + wb.Sheet1.A3());
        // outputs 42
        console.log('result: ' + wb.result());
        // Outputs 42
      });

# Author

Written by Michael Zedeler <michael@zedeler.dk>.

# License

See the LICENSE file.
