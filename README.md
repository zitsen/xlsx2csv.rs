# xlsx2csv - An Excel-like spreadsheet to CSV coverter writen in Rust.

```
USAGE:
    xlsx2csv [FLAGS] [OPTIONS] <xlsx> [output]...

FLAGS:
    -h, --help               Prints help information
    -i, --ignore-case        Rgex case insensitivedly
    -l, --list               List sheet names by id
    -u, --use-sheet-names    Use sheet names as output filename prefix (in current dir or --workdir)
    -V, --version            Prints version information

OPTIONS:
    -d, --delimiter <delimiter>    Delimiter for output [default: ,]
    -X, --exclude <exclude>        A regex pattern for matching sheetnames to exclude, used with '-u'
    -I, --include <include>        A regex pattern for matching sheetnames to include, used with '-u'
    -s, --select <select>          Select sheet by name or id in output, only used when output to stdout
    -w, --workdir <workdir>        Output files location if `--use-sheet-names` setted

ARGS:
    <xlsx>         Input Excel-like files, supports: .xls .xlsx .xlsb .xlsm .ods
    <output>...    Output each sheet to sperated file
```

## Install

```sh
cargo install xlsx2csv
```

## Advanced Usage

### output sheets one-by-one

Simple usage is similar to `ssconvert` syntax, like this:

```sh
xlsx2csv input.xlsx sheet1.csv sheet2.csv
```

This will output the first to `sheet1.csv`, the second to `sheet2.csv`, and ignore other sheets.

### pipe output

If no output position args setted, eg. `xlsx2csv input.xlsx`, it'll write first sheet to stdout. So the two commands are equal:

- `xlsx2csv input.xlsx sheet1.csv`
- `xlsx2csv input.xlsx > sheet1.csv`.

If you want to select specific sheet to stdout, use `-s/--select <id or name>` (id is 0-based):

```sh
xlsx2csv input.xlsx -s 1
```

In previous command, it'll output the second(0-based 1 is the second) sheet to stdout.

### list sheetnames

Use `--list/-l` it will just print all the sheetnames by id.

```sh
xlsx2csv --list
xlsx2csv -l
```

### multiple sheets without filename setted

If there's many sheets that you don't wanna set filename for each,
use `-u/--use-sheet-names` to write with sheetnames.

```sh
xlsx2csv input.xlsx -u
```

If you want to write to directory other than `.`, use `-w/--workdir` along with `-u` option.

```
xlsx2csv input.xlsx -u -w test/
```

The filename extension is detemined by delimiter, `,` to `.csv`, `\t` to `.tsv`, others will treat as ','.

### multiple sheets matching or not matching a regex pattern

By default, it will output all sheets, but if you want to select by sheet names with regex match, use `-I/--include` to include only matching, and `-X/--exclude` to exclude matching.
You could also combine these two option with *include-first-exclude-after* order:

```sh
xlsx2csv input.xlsx -I 'a\d+'
xlsx2csv input.xlsx -X '\s'
xlsx2csv input.xlsx -I '\S{3,}' -X 'Sheet'
```

The last command line will first include all sheet with pattern '\S{3,}' matched and then exclude that match `Sheet`.

## Detailed options

The following is printed by `xlsx2csv -h`

```
USAGE:
    xlsx2csv [FLAGS] [OPTIONS] <xlsx> [output]...

FLAGS:
    -h, --help               
            Prints help information

    -l, --list               
            List sheet names by id

    -u, --use-sheet-names    
            Use sheet names as output filename prefix (in current dir or --workdir)

    -V, --version            
            Prints version information


OPTIONS:
    -d, --delimiter <delimiter>    
            Delimiter for output.
            
            If `use-sheet-names` setted, it will control the output filename extension: , -> csv, \t -> tsv [default: ,]
    -s, --select <select>          
            Select sheet by name or id in output, only used when output to stdout

    -w, --workdir <workdir>        
            Output files location if --use-sheet-names setted


ARGS:
    <xlsx>         
            Input Excel-like files, supports: .xls .xlsx .xlsb .xlsm .ods

    <output>...    
            Output each sheet to sperated file.
            
            If not setted, output first sheet to stdout.
```

## License

[MIT](LICENCE-MIT) OR [Apache-2.0](LICENCE-APACHE)
