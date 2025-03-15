use calamine::{Reader, Xlsx, open_workbook};
use std::error::Error;
use xlsxwriter::{Workbook, Worksheet};
use indexmap::IndexMap; // Add this import for ordered map

// Define an enum for the match states
enum MatchState {
    NoMatch,       // No match found
    PartialMatch,  // Similar but not exact match
    ExactMatch     // Exact match found
}

fn main() -> Result<(), Box<dyn Error>> {
    // File paths
    let file_a_path = "珠海金灶储能电站点表.xlsx"; // TODO: change target workbook name
    let file_b_path = "南沙丰田-校对.xlsx";
    let output_path = "a_compared.xlsx";

    println!("Starting comparison...");

    // Use IndexMap instead of HashMap to preserve insertion order
    let mut b_file_map: IndexMap<String, String> = IndexMap::new();

    // Open file B
    println!("Reading file B: {}", file_b_path);
    let mut workbook_b: Xlsx<_> = open_workbook(file_b_path)?;
    let range_b = workbook_b
        .worksheet_range("编码修编")
        .map_err(|_| "Could not find 编码修编 in file B")?;

    // Populate IndexMap from file B
    for row in range_b.rows().skip(9) {
        // Skip header row
        // skip none data rows
        if row.len() >= 17 && row[33].to_string() != "" {
            let key = row[33].to_string();
            let value = row[34].to_string();
            b_file_map.insert(key, value);
        }
    }

    // Open file A
    println!("Reading file A: {}", file_a_path);
    let mut workbook_a: Xlsx<_> = open_workbook(file_a_path)?;
    let range_a = workbook_a
        .worksheet_range("珠海金灶储能电站点表") // TODO: change target sheet name
        .map_err(|_| "Could not find 珠海金灶储能电站点表 in file A")?;

    // Create output workbook
    let workbook_out = Workbook::new(output_path)?;
    let mut sheet_out = workbook_out.add_worksheet(Some("Sheet1"))?;

    // Write headers
    write_cell(&mut sheet_out, 0, 0, "结果")?;
    write_cell(&mut sheet_out, 0, 1, "备注")?;

    // Collect the first column of each row in range_a
    let a_row: Vec<String> = range_a
        .rows()
        .skip(1)
        .map(|row| row[1].to_string()) // TODO: change column index
        .collect();

    // Use IndexMap for results to maintain insertion order
    let mut result: IndexMap<String, String> = IndexMap::new();
    
    for (key, value) in b_file_map {
        let mut match_state = MatchState::NoMatch;
        let mut similar: Vec<String> = Vec::new();
        result.insert(key.clone(), "".to_string());

        // find similar item collection
        for row in &a_row {
            if row.contains(&key) {
                match_state = MatchState::PartialMatch;
                similar.push(row.clone());
            }
        }
        
        if !similar.is_empty() {
            for i in similar {
                if compare_strings(&value, &i) {
                    match_state = MatchState::ExactMatch;
                }
            }
        }

        match match_state {
            MatchState::NoMatch => {
                result.insert(key, "点位缺失".to_string());
            },
            MatchState::PartialMatch => {
                result.insert(key, "异常".to_string());
            },
            MatchState::ExactMatch => {
                result.insert(key, "正常".to_string());
            }
        }
    }

    // Write result to file - order will be preserved because we're using IndexMap
    for (i, (key, value)) in result.iter().enumerate() {
        write_cell(&mut sheet_out, (i + 1) as u32, 0, key)?;
        write_cell(&mut sheet_out, (i + 1) as u32, 1, value)?;
    }

    // Format the columns for better readability
    sheet_out.set_column(0, 0, 20.0, None)?;
    sheet_out.set_column(1, 1, 20.0, None)?;
    sheet_out.set_column(2, 2, 30.0, None)?;

    workbook_out.close()?;

    println!("Comparison complete!");
    println!("Results saved to: {}", output_path);

    Ok(())
}

// Helper function to write a cell value to the output worksheet
fn write_cell(
    sheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: &str,
) -> Result<(), Box<dyn Error>> {
    // Try to parse as number first
    if let Ok(num) = value.parse::<f64>() {
        sheet.write_number(row, col, num, None)?;
    } else {
        sheet.write_string(row, col, value, None)?;
    }
    Ok(())
}

fn compare_strings(s1: &str, s2: &str) -> bool {
    // Define the prefixes to check
    let prefix1 = "DD_SSNS_S";
    let prefix2 = "DD_ZHJZ_S"; // TODO: alert prefix

    // Extract the parts after the prefixes
    let part1 = s1.strip_prefix(prefix1).unwrap_or(s1);
    let part2 = s2.strip_prefix(prefix2).unwrap_or(s2);
    
    // Compare the remaining parts
    part1 == part2
}