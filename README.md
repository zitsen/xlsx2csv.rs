# xlsx2csv

xlsx2csv - Excel XLSX to CSV coverter in Rust

## Usage:

```text
xlsx2csv 0.1.0
Huo Linhe <linhehuo@gmail.com>
Excel XLSX to CSV converter

USAGE:
    xlsx2csv [FLAGS] [OPTIONS] --xlsx <xlsx>

    FLAGS:
        -h, --help           Prints help information
        -S, --sheet_names    List sheet names if you want to use --sheets option
        -V, --version        Prints version information

    OPTIONS:
        -d, --delimiter <delimiter>    The field delimiter for reading CSV data.
                                       [default: ,]
        -o, --directory <directory>    Output directory [default: .]
        -s, --sheet <NAME>...          Select sheets
        -x, --xlsx <xlsx>              Excel file with XLSX format

```
## Installation

`cargo install xslx2csv`

To build from source:

```
git clone https://github.com/zitsen/xlsx2csv.rs
cd xlsx2csv.rs
cargo install .
```

## Examples

Use it simply, and convert each worksheet to csv file in current directory

```zsh
xlsx2csv -x test.xlsx
```

Choose to convert some of worksheets

```zsh
xlsx2csv -S -x test.xlsx
xlsx2csv -x test.xlsx -s sheet1 -s sheet2
```

Output settings:

```zsh
xlsx2csv ... --directory /tmp --delimiter '\t'
```


License: MIT/Apache-2.0
