# simple_excel_writer
simple excel writer in Rust

[![Build Status](https://travis-ci.org/outersky/simple_excel_writer.png?branch=master)](https://travis-ci.org/outersky/simple_excel_writer) 
[Documentation](https://docs.rs/simple_excel_writer/)

## Example

```rust,no_run
extern crate simple_excel_writer as excel;

use excel::*;

fn main() {
    let mut wb = Workbook::create("/tmp/b.xlsx");
    let mut sheet = wb.create_sheet("SheetName");

    // set column width
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 80.0 });
    sheet.add_column(Column { width: 60.0 });

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["Name", "Title","Success","XML Remark"])?;
        sw.append_row(row!["Amy", (), true,"<xml><tag>\"Hello\" & 'World'</tag></xml>"])?;
        sw.append_blank_rows(2);
        sw.append_row(row!["Tony", blank!(2), "retired"])
    }).expect("write excel error!");

    let mut sheet = wb.create_sheet("Sheet2");
    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["Name", "Title","Success","Remark"])?;
        sw.append_row(row!["Amy", "Manager", true])
    }).expect("write excel error!");

    wb.close().expect("close excel error!");
}
```

## Todo

- support style

## Change Log

#### 0.2.0 (2019-06-11)
- basic support for outputting formulas

#### 0.1.4 (2017-03-24)
- escape xml characters.

#### 0.1.3 (2017-01-03)
- support 26+ columns .
- fix column width bug.

#### 0.1.2 (2017-01-02)
- support multiple sheets

#### 0.1 (2017-01-01)
- generate the basic xlsx file

## License
Apache-2.0
