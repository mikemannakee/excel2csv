use calamine::{open_workbook_auto, DataType, Reader};
use std::{env, io::Write};
use regex::Regex;

fn main() {
	let mut delim = ",";  // Default delimiter is a comma
	let args: Vec<String> = env::args().collect();

	#[cfg(debug_assertions)]
	dbg!(&args);

	// Check if the user has provided the filename
	let filename = args[1].clone();
	let len = args.len();
	let out_filename: String = match len {
		2 => match filename.rfind(".") { // If the user has not provided the output filename, we will use the input filename and change the extension to .csv
			Some(i) => format!("{}{}", &filename[..i], ".csv"),
			None => format!("{}.csv", filename),
		},
		_ => args[2].clone(), // If the user has provided the output filename, we will use that
	};

	// Check if the user has provided the delimiter
	if len > 3 {
		delim = args[3].as_str();
	}

	#[cfg(debug_assertions)]
	dbg!("Out filename: {}", &out_filename);

	// Check if the file exists
	let exists = std::path::Path::new(&filename).exists();
	if !exists {
		println!("Error: File '{}' does not exist", filename);
		return;
	}

	// Open the workbook
	let mut workbook = open_workbook_auto(&filename).expect("Cannot open file");

	let sheets = workbook.sheet_names().to_owned();

	// Regex to remove non-printable characters
	let re = Regex::new(r"[^\x20-\x7E]").unwrap();

	#[cfg(debug_assertions)]
	dbg!(&sheets);

	let mut out = String::new();
	for sheet in sheets.iter() {
		if let Some(Ok(range)) = workbook.worksheet_range(sheet) {
			for row in range.rows() {
				// A couple of flags to determine if the row is empty
				let mut empty = true;
				let mut line = String::new();
				for cell in row.iter() {
					match cell {
						DataType::Empty => line.push_str(delim),
						DataType::String(s) => {
							// print!("'{}' is a string! \n", s);
							let only_printable = re.replace_all(s, ""); // Remove non-printable characters
							line.push_str(&format!("{}{}", only_printable.replace(delim, &format!("\\{}", delim)), delim)); // Escape the delimiter
							empty = false;
						}
						DataType::Float(f) => {
							line.push_str(&format!("{}{}", f, delim));
							empty = false;
						}
						DataType::Int(i) => {
							line.push_str(&format!("{}{}", i, delim));
							empty = false;
						}
						DataType::Bool(b) => {
							line.push_str(&format!("{}{}", b, delim));
							empty = false;
						}
						DataType::Error(e) => {
							line.push_str(&format!("{}{}", e, delim));
							empty = false;
						}
						DataType::DateTime(dt) => {
							line.push_str(&format!("{}{}", dt, delim));
							empty = false;
						}
						DataType::Duration(d) => {
							line.push_str(&format!("{}{}", d, delim));
							empty = false;
						}
						DataType::DateTimeIso(dt) => {
							line.push_str(&format!(
								"{}{}",
								dt.replace(delim, &format!("\\{}", delim)),
								delim
							));
							empty = false;
						}
						DataType::DurationIso(d) => {
							line.push_str(&format!(
								"{}{}",
								d.replace(delim, &format!("\\{}", delim)),
								delim
							));
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

	// Write the output to a file
	let mut out_file = std::fs::File::create(&out_filename).expect("Cannot create file");
	out_file
		.write_all(out.as_bytes())
		.expect("Cannot write to file");

	// A message letting the user know that the file has been converted
	println!(
		"File '{}' converted to '{}' using '{}' as the delimiter",
		filename, out_filename, delim
	);
}
