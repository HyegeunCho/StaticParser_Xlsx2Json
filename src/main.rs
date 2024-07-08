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

/// Hashmap 구조
/// {sheet_name: [{ID: 0, KEY_A: 1, KEY_B: 2}, ...], sheet_name2: [...]}

fn serialize_to_json(hash_map: &HashMap<String, Vec<HashMap<String, String>>>) -> Result<(), Box<dyn Error>> {

    let mut map = HashMap::new();
    map.insert("key1", "value1");
    map.insert("key2", "value2");
    map.insert("key3", "value3");

    // HashMap을 JSON 문자열로 변환
    let json = serde_json::to_string_pretty(&hash_map).expect("Failed to convert HashMap to JSON");

    // JSON 문자열 출력
    // println!("{}", json);

    let mut file = File::create("bd_static.json")?;
    file.write_all(json.as_bytes())?;
    
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let path = "bd_static.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    let mut static_data: HashMap<String, Vec<HashMap<String, String>>> = HashMap::new();
    // print sheet names
    let mut sheet_index: i64 = 0;
    for sheet_name in workbook.sheet_names().to_owned() {
        sheet_index = sheet_index + 1;
        if sheet_name.starts_with('_') {
            continue;
        }
        // println!("[{}] {}", sheet_index, sheet_name);

        
        let mut sheet_data: Vec<HashMap<String, String>> = Vec::new();
        let range_result = workbook.worksheet_range(&sheet_name);
        match range_result {
            Ok(range) => {

                let header_from_fn = get_headers_from_sheet(&range);
                let mut header_keys: Vec<i64> = header_from_fn.keys().map(|i| i.clone()).collect();
                header_keys.sort();

                let mut row_index: i64 = 0;
                for row in range.rows() {
                    
                    let mut row_data: HashMap<String, String> = HashMap::new();
                    if row_index > 0 {
                        for hIndex in &header_keys {
                            let header_name = header_from_fn.get(&hIndex).unwrap_or(&("".to_string())).clone();
                            let cell_value = cell_data_to_string(row.get(*hIndex as usize).ok_or("".to_string()).unwrap());
                            row_data.insert(header_name.clone(), cell_value.clone());
                        }
                        sheet_data.push(row_data);
                    }
                    row_index = row_index + 1;
                }
            },
            Err(e) => {
                println!("에러 발생");
            }
        }
        
        static_data.insert(sheet_name.clone(), sheet_data);
        serialize_to_json(&static_data);
    }

    
    Ok(())
}
