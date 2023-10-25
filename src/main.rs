use std::{env, io::Write};
use calamine::{Reader, open_workbook_auto, DataType};

fn main() {
	let mut delim = ",";
	let args: Vec<String> = env::args().collect();

	#[cfg(not(release))]
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

	#[cfg(not(release))]
	dbg!("Out filename: {}", &out_filename);

	let exists = std::path::Path::new(&filename).exists();
	if !exists {
		println!("Error: File '{}' does not exist", filename);
		return;
	}

	let mut workbook = open_workbook_auto(&filename).expect("Cannot open file");

	let sheets = workbook.sheet_names().to_owned();

	#[cfg(not(release))]
	dbg!(&sheets);
	
	let mut out = String::new();
	for sheet in sheets.iter() {
		if let Some(Ok(range)) = workbook.worksheet_range(sheet) {
			for row in range.rows() {
				for cell in row.iter() {
					match cell {
						DataType::Empty => out.push_str(delim),
						DataType::String(s) => out.push_str(&format!("{}{}", s, delim).replace(delim, "")),
						DataType::Float(f) => out.push_str(&format!("{}{}", f, delim).replace(delim, "")),
						DataType::Int(i) => out.push_str(&format!("{}{}", i, delim).replace(delim, "")),
						DataType::Bool(b) => out.push_str(&format!("{}{}", b, delim).replace(delim, "")),
						DataType::Error(e) => out.push_str(&format!("{}{}", e, delim).replace(delim, "")),
						DataType::DateTime(dt) => out.push_str(&format!("{}{}", dt, delim).replace(delim, "")),
						DataType::Duration(d) => out.push_str(&format!("{}{}", d, delim).replace(delim, "")),
						DataType::DateTimeIso(dt) => out.push_str(&format!("{}{}", dt, delim).replace(delim, "")),
						DataType::DurationIso(d) => out.push_str(&format!("{}{}", d, delim).replace(delim, ""))
					}
				}
				out.push_str("\n");
			}
		}
	}
	
	let mut out_file = std::fs::File::create(&out_filename).expect("Cannot create file");
	out_file.write_all(out.as_bytes()).expect("Cannot write to file");

}
