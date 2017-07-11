extern crate calamine;
extern crate csv;
#[macro_use]
extern crate clap;
extern crate pbr;

use std::path::Path;

use clap::{Arg, App};
use calamine::Sheets;
use calamine::DataType;
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

    if matches.is_present("sheet_names") {
        let mut workbook = Sheets::open(xlsx).expect("open xlsx file");
        for (i, sheet) in workbook.sheet_names().unwrap().iter().enumerate() {
            println!("{}\t{}", i, sheet);
        }
        return;
    }
    let mut workbook = Sheets::open(xlsx).expect("open xlsx file");

    let sheets: Vec<String>= matches.values_of("sheet")
        .map(|sheet| sheet.map(|s| s.to_string()).collect())
        .unwrap_or(workbook.sheet_names().unwrap());

    let output = matches.value_of("directory").unwrap();
    let delimiter = matches.value_of("delimiter").unwrap().as_bytes().first().unwrap();

    for sheet in sheets {
        let path = Path::new(output).join(format!("{}.csv", sheet));
        println!("* prepring write to {}", path.display());
        let range = workbook.worksheet_range(&sheet).expect(&format!("find sheet {}", sheet));
        let mut wtr = csv::WriterBuilder::new().delimiter(*delimiter).from_path(path).expect("open csv");
        let size = range.get_size();
        println!("** sheet range size is {:?}", size);
        println!("** start writing", );
        let mut pb = ProgressBar::new(100);
        let rows = range.rows();
        for (i, row) in rows.enumerate() {
            if i % (size.0 / 100) == 0 {
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
