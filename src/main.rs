//! xlsx2csv - Excel XLSX to CSV coverter in Rust
//!
//! # Usage:
//!
//! ```text
//! xlsx2csv
//! Huo Linhe <linhehuo@gmail.com>
//! Excel XLSX to CSV converter
//!
//! USAGE:
//!     xlsx2csv [FLAGS] [OPTIONS] --xlsx <xlsx>
//!
//!     FLAGS:
//!         -h, --help           Prints help information
//!         -S, --sheet_names    List sheet names if you want to use --sheets option
//!         -V, --version        Prints version information
//!
//!     OPTIONS:
//!         -d, --delimiter <delimiter>    The field delimiter for reading CSV data.
//!                                        [default: ,]
//!         -o, --directory <directory>    Output directory [default: .]
//!         -s, --sheet <NAME>...          Select sheets by name
//!         -r, --replace <PREFIX>...      Replace output csv filename, default is sheet name
//!         -x, --xlsx <xlsx>              Excel file with XLSX format
//!
//! ```
//!
//! # Examples
//!
//! Use it simply, and convert each worksheet to csv file in current directory
//!
//! ```zsh
//! xlsx2csv -x test.xlsx
//! ```
//!
//! Choose to convert some of worksheets
//!
//! ```zsh
//! # get sheet names
//! xlsx2csv -S -x test.xlsx
//! #0,sheet1
//! #1,sheet2
//! xlsx2csv -x test.xlsx -s sheet1 sheet2
//! #sheet1.csv
//! #sheet2.csv
//! xlsx2csv -x test.xlsx -s sheet1 sheet2 -r foo bar
//! #foo.csv
//! #bar.csv
//! ```
//!
//! Output settings:
//!
//! ```zsh
//! xlsx2csv ... --directory /tmp --delimiter '\t'
//! ```
//!
extern crate calamine;
extern crate csv;
#[macro_use]
extern crate clap;
extern crate pbr;

use std::path::Path;

use calamine::{open_workbook, Xlsx};
use calamine::Reader;
use calamine::DataType;
use clap::{Arg, App};
use pbr::ProgressBar;

fn main() {
    let matches = App::new(crate_name!())
        .about(crate_description!())
        .version(crate_version!())
        .author(crate_authors!())
        .arg(Arg::with_name("xlsx")
             .short("x")
             .long("xlsx")
             .help("Excel file with XLSX format")
             .takes_value(true)
             .required(true))
        .arg(Arg::with_name("sheet_names")
             .short("S")
             .long("sheet_names")
             .help("List sheet names if you want to use --sheets option")
             .conflicts_with_all(&["sheets"]))
        .arg(Arg::with_name("sheet")
             .short("s")
             .long("sheet")
             .value_name("NAME")
             .help("Select sheets")
             .takes_value(true)
             .multiple(true))
        .arg(Arg::with_name("replace")
             .short("r")
             .long("replace")
             .value_name("filename")
             .help("replace selected sheet names")
             .takes_value(true)
             .multiple(true))
        .arg(Arg::with_name("directory")
             .short("o")
             .long("directory")
             .help("Output directory")
             .takes_value(true)
             .default_value("."))
        .arg(Arg::with_name("delimiter")
             .short("d")
             .long("delimiter")
             .help("The field delimiter for reading CSV data.")
             .default_value(",")
             .takes_value(true))
        .get_matches();
    let xlsx = matches.value_of("xlsx").unwrap();
    let mut workbook: Xlsx<_> = open_workbook(xlsx).expect("open xlsx file");
    let sheets = workbook.sheet_names();
    
    if matches.is_present("sheet_names") {
        for (i, sheet) in workbook.sheet_names().iter().enumerate() {
            println!("{}\t{}", i, sheet);
        }
        return;
    }

    let sheets: Vec<String>= matches.values_of("sheet")
        .map(|sheet| sheet.map(|s| s.to_string()).collect())
        .unwrap_or(sheets.into_iter().map(Clone::clone).collect());
    let replaces: Vec<String> = matches.values_of("replace")
        .map(|sheet| sheet.map(|s| s.to_string()).collect())
        .unwrap_or(sheets.clone());

    assert_eq!(sheets.len(), replaces.len(), "sheets number must be equal to replaces");

    let output = matches.value_of("directory").unwrap();
    let delimiter = matches.value_of("delimiter").unwrap().as_bytes().first().unwrap();

    for (sheet, replace) in sheets.iter().zip(replaces.iter()) {
        let path = Path::new(output).join(format!("{}.csv", replace));
        println!("* prepring write to {}", path.display());
        let range = workbook.worksheet_range(&sheet).expect(&format!("find sheet {}", sheet)).expect("get range");
        let mut wtr = csv::WriterBuilder::new().delimiter(*delimiter).from_path(path).expect("open csv");
        let size = range.get_size();
        println!("** sheet range size is {:?}", size);
        if size.0 == 0 || size.1 == 0 {
            eprintln!("worsheet range sizes should not be 0, continue");
            continue;
        }
        println!("** start writing", );
        let mut pb = ProgressBar::new(if size.0 > 100 { 100 } else { size.0 as _ });
        let rows = range.rows();
        for (i, row) in rows.enumerate() {
            if size.0 <=100 {
                pb.inc();
            } else if i % (size.0 / 100) == 0 {
                pb.inc();
            }
            let cols: Vec<String> = row.iter().map(|c| {
                match *c {
                    DataType::Int(ref c) => format!("{}", c),
                    DataType::Float(ref c) => format!("{}", c),
                    DataType::String(ref c) => format!("{}", c),
                    DataType::Bool(ref c) => format!("{}", c),
                    _ => "".to_string(),
                }
            }).collect();
            wtr.write_record(&cols).unwrap();
        }
        pb.finish_print("** done, flush to write csv file");
        wtr.flush().unwrap();
    }
}
