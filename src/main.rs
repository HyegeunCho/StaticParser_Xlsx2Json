use calamine::{open_workbook, Data, DataType, Range, RangeDeserializerBuilder, Reader, Xlsx};
use serde::Serialize;
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::Write;
use std::ops::Index;
use std::path::is_separator;
use calamine::Data::Int;
// #[derive(Serialize)]
// struct Sheet {
//     name: String,
//     data: Vec<HashMap<String, dyn DataType>>,
// }

fn cell_data_to_string(data: &Data) -> String {
    let cell_value = match data {
        calamine::Data::Empty => "".to_string(),
        calamine::Data::String(s) => s.clone(),
        calamine::Data::Float(f) => f.to_string(),
        calamine::Data::Int(i) => i.to_string(),
        calamine::Data::Bool(b) => b.to_string(),
        calamine::Data::Error(e) => "".to_string(),
        calamine::Data::DateTime(d) => d.to_string(),
        calamine::Data::DateTimeIso(diso) => diso.to_string(),
        calamine::Data::DurationIso(duiso) => duiso.to_string()
    };
    cell_value
}

fn get_headers_from_sheet(sheet_range: &calamine::Range<Data>) -> HashMap<i64, String> {

    let mut result: HashMap<i64, String> = HashMap::new();

    let mut cell_index: i64 = 0;


    let mut rows = sheet_range.rows();
    if let Some(first_row) = rows.next() {

        for cell in first_row {
            let cell_value = cell_data_to_string(cell);
            if !cell_value.starts_with('_') {
                result.insert(cell_index, cell_value.clone());
            }
            cell_index = cell_index + 1;
        }
    }

    result.remove(&(cell_index - 1));
    result
} 

fn main() -> Result<(), Box<dyn Error>> {
    let path = "bd_static.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // print sheet names
    let mut sheet_index: i64 = 0;
    for sheet_name in workbook.sheet_names().to_owned() {
        sheet_index = sheet_index + 1;
        if sheet_name.starts_with('_') {
            continue;
        }
        println!("[{}] {}", sheet_index, sheet_name);

        if sheet_index == 4 {
            let range_result = workbook.worksheet_range(&sheet_name);

            let mut header: HashMap<i64, String> = HashMap::new();
            match range_result {
                Ok(range) => {

                    let header_from_fn = get_headers_from_sheet(&range);
                    let mut header_keys: Vec<i64> = header_from_fn.keys().map(|i| i.clone()).collect();
                    header_keys.sort();

                    let mut row_index: i64 = 0;
                    for row in range.rows() {
                        
                        if row_index > 0 {

                            for hIndex in &header_keys {
                        
                                let header_name = header_from_fn.get(&hIndex).unwrap_or(&("".to_string())).clone();
                                let cell_value = cell_data_to_string(row.get(*hIndex as usize).ok_or("".to_string()).unwrap());
                                println!("[{:?} : {:?}]", header_name, cell_value);
                                

                                // if let Some(cell_value) = row.get(*hIndex as usize).map(|c| cell_data_to_string(c)) {
                                //     println!("[{:?}] {} {:?}", header_name, hIndex, cell_value);
                                // }
                                // else{
                                    
                                // }
                            }

                        }
                        row_index = row_index + 1;
                    }
                },
                Err(e) => {
                    println!("에러 발생");
                }
            }
        }
    }

    



    // for sheet_name in workbook.sheet_names().to_owned() {
        
    //     if sheet_name.starts_with('_') {
    //         continue;  
    //     }

    //     let range_result = workbook.worksheet_range(&sheet_name);
    //     match  range_result {
    //         Ok(range) => {
    //             let mut index : i64 = 0;
    //             for row in range.rows() {
    //                 for cell in row {
    //                     println!("{:?}", cell.get_string());
    //                 }
    //                 index = index + 1;
    //             }
                
    //             // let mut rangeIter = RangeDeserializerBuilder::new().from_range(&range)?;
    //             // if let Some(result) = rangeIter.next() {
    //             //     let (label, value): (String, f64) = result?;
    //             // }
    //         },
    //         Err(e) => {
                
    //         }
    //     }
        
    //     // let range = workbook.worksheet_range(&sheet_name)
    //     //     .ok_or(Error::Msg("Cannot find " + &sheet_name))??;
    //     // 
    //     // let mut rangeIter = RangeDeserializerBuilder::new().from_range(&range)?;
    //     // if let Some(result) = rangeIter.next() {
    //     //     let (label, value): (String, f64) = result?;
    //     //     assert_eq!(label, "celsius");
    //     //     assert_eq!(value, 22.2222);
    //     //     Ok(());
    //     // } 
         
    //     // if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
    //     //     let mut rows = range.rows();
    //     // 
    //     //     if let Some(headers) = rows.next() {
    //     //         let headers: Vec<String> = headers.iter().map(|h| h.to_string()).collect();
    //     // 
    //     //         let mut sheet_data = Vec::new();
    //     // 
    //     //         for row in rows {
    //     //             let mut row_data = HashMap::new();
    //     //             for (header, cell) in headers.iter().zip(row.iter()) {
    //     //                 row_data.insert(header.clone(), cell.clone());
    //     //             }
    //     //             sheet_data.push(row_data);
    //     //         }
    //     // 
    //     //         // sheets.push(Sheet {
    //     //         //     name: sheet_name.clone(),
    //     //         //     data: sheet_data,
    //     //         // });
    //     //     }
    //     // }
    // }

    // let json = serde_json::to_string_pretty(&sheets)?;
    // let mut file = File::create("output.json")?;
    // file.write_all(json.as_bytes())?;

    Ok(())
}
