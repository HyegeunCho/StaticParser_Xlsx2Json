use calamine::{open_workbook, DataType, Reader, Xlsx, RangeDeserializerBuilder};
use serde::Serialize;
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::Write;
use std::path::is_separator;
use calamine::Data::Int;
// #[derive(Serialize)]
// struct Sheet {
//     name: String,
//     data: Vec<HashMap<String, dyn DataType>>,
// }

fn main() -> Result<(), Box<dyn Error>> {
    let path = "bd_static.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // let mut sheets = Vec::new();

    for sheet_name in workbook.sheet_names().to_owned() {
        
        if sheet_name.starts_with('_') {
            continue;  
        }

        let range_result = workbook.worksheet_range(&sheet_name);
        match  range_result {
            Ok(range) => {
                let mut index : i64 = 0;
                for row in range.rows() {
                    for cell in row {
                        println!("{:?}", cell.get_string());
                    }
                    index = index + 1;
                }
                
                let mut rangeIter = RangeDeserializerBuilder::new().from_range(&range)?;
                if let Some(result) = rangeIter.next() {
                    let (label, value): (String, f64) = result?;
                    assert_eq!(label, "celsius");
                    assert_eq!(value, 22.2222);
                }
            },
            Err(e) => {
                
            }
        }
        
        // let range = workbook.worksheet_range(&sheet_name)
        //     .ok_or(Error::Msg("Cannot find " + &sheet_name))??;
        // 
        // let mut rangeIter = RangeDeserializerBuilder::new().from_range(&range)?;
        // if let Some(result) = rangeIter.next() {
        //     let (label, value): (String, f64) = result?;
        //     assert_eq!(label, "celsius");
        //     assert_eq!(value, 22.2222);
        //     Ok(());
        // } 
         
        // if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
        //     let mut rows = range.rows();
        // 
        //     if let Some(headers) = rows.next() {
        //         let headers: Vec<String> = headers.iter().map(|h| h.to_string()).collect();
        // 
        //         let mut sheet_data = Vec::new();
        // 
        //         for row in rows {
        //             let mut row_data = HashMap::new();
        //             for (header, cell) in headers.iter().zip(row.iter()) {
        //                 row_data.insert(header.clone(), cell.clone());
        //             }
        //             sheet_data.push(row_data);
        //         }
        // 
        //         // sheets.push(Sheet {
        //         //     name: sheet_name.clone(),
        //         //     data: sheet_data,
        //         // });
        //     }
        // }
    }

    // let json = serde_json::to_string_pretty(&sheets)?;
    // let mut file = File::create("output.json")?;
    // file.write_all(json.as_bytes())?;

    Ok(())
}
