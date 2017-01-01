# simple_excel_writer
simple excel writer in Rust

##Example

```rust,no_run
    extern crate simple_excel_writer as excel;
    
    use excel::*;
    
    fn main() {
        println!("hello");
    
        let mut wb = Workbook::create("/tmp/b.xlsx");
        let mut sheet = wb.create_sheet("Sheet2");
    
    /*
        // optional: set column width
        sheet.add_column(Column { custom_width: 1.0, width: 50.0 });
        sheet.add_column(Column { custom_width: 1.0, width: 50.0 });
        sheet.add_column(Column { custom_width: 1.0, width: 150.0 });
    */
    
        wb.write_sheet(&mut sheet, |sheet_writer| {
            let mut r = Row::new();
            r.add_cell(1, true);
            r.add_cell(2, "Hello");
            sheet_writer.append_row(r)?;
    
            sheet_writer.append_empty_rows(5);
    
            r = Row::new();
            r.add_cell(1, 3.6);
            r.add_cell(2, "World");
            r.add_cell(3, "!");
            sheet_writer.append_row(r)
        }).expect("write excel error!");
    
        wb.close().expect("close excel error!");
    }
```

##Note

Only tested on macOS for the time being, because it depends on locale `zip` command ( `zip -q -r target.xlsx excel_tmp_dir/*`).

##Todo

- use `zip` crate to compress files to xlsx. (currently error occurs while opening xlsx file generated by `zip`)

##Change Log
####0.1 (2017-01-01)
- generate the basic xlsx file