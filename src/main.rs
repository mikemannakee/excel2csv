use std::{env, io::Write};
use calamine::{Reader, open_workbook_auto, DataType};
use regex::Regex;

fn main() {
	let mut delim = ",";
	let args: Vec<String> = env::args().collect();

	#[cfg(debug_assertions)]
	dbg!(&args);

	let filename = args[1].clone();
	let len = args.len();
	let out_filename: String = match len {
		2 => match filename.rfind(".") {
			Some(i) => format!("{}{}", &filename[..i], ".csv"),
			None => format!("{}.csv", filename)
		},
		_ => args[2].clone()
	};

	if len > 3 {
		delim = args[3].as_str();
	}

	#[cfg(debug_assertions)]
	dbg!("Out filename: {}", &out_filename);

	let exists = std::path::Path::new(&filename).exists();
	if !exists {
		println!("Error: File '{}' does not exist", filename);
		return;
	}

	let mut workbook = open_workbook_auto(&filename).expect("Cannot open file");

	let sheets = workbook.sheet_names().to_owned();

	#[cfg(debug_assertions)]
	dbg!(&sheets);
	
	let mut out = String::new();
	let mut empty = false;
	let mut line = String::new();
	for sheet in sheets.iter() {
		if let Some(Ok(range)) = workbook.worksheet_range(sheet) {
			for row in range.rows() {
				// A couple of flags to determine if the row is empty
				empty = true;
				line = "".to_string();
				for cell in row.iter() {
					match cell {
						DataType::Empty => line.push_str(delim),
						DataType::String(s) => { 
							line.push_str(&format!("{}{}", s, delim).replace(delim, ""));
							empty = false;
						},
						DataType::Float(f) => {
							line.push_str(&format!("{}{}", f, delim).replace(delim, ""));
							empty = false;
						},
						DataType::Int(i) => {
							line.push_str(&format!("{}{}", i, delim).replace(delim, ""));
							empty = false;
						},
						DataType::Bool(b) => {
							line.push_str(&format!("{}{}", b, delim).replace(delim, ""));
							empty = false;
						},
						DataType::Error(e) => {
							line.push_str(&format!("{}{}", e, delim).replace(delim, ""));
							empty = false;
						},
						DataType::DateTime(dt) => {
							line.push_str(&format!("{}{}", dt, delim).replace(delim, ""));
							empty = false;
						},
						DataType::Duration(d) => {
							line.push_str(&format!("{}{}", d, delim).replace(delim, ""));
							empty = false;
						},
						DataType::DateTimeIso(dt) => {
							line.push_str(&format!("{}{}", dt, delim).replace(delim, ""));
							empty = false;
						},
						DataType::DurationIso(d) => {
							line.push_str(&format!("{}{}", d, delim).replace(delim, ""));
							empty = false;
						}
					}
				}
				if !empty {
					out.push_str(&line);
					out.push_str("\n");
				}
			}
		}
	}

	
	let mut out_file = std::fs::File::create(&out_filename).expect("Cannot create file");
	out_file.write_all(out.as_bytes()).expect("Cannot write to file");

	println!("File '{}' converted to '{}' using '{}' as the delimiter", filename, out_filename, delim);
}
