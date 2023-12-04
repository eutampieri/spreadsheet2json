use calamine::Reader;
use clap::Parser;
use clap::ValueEnum;
use std::fmt::Display;

#[derive(Debug, Clone, ValueEnum)]
enum Format {
    /// The output will consist of an array of objects
    Object,
    /// The output will be a matrix
    List,
}

impl Display for Format {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(
            format_args!("{:?}", self)
                .to_string()
                .to_lowercase()
                .as_str(),
        )
    }
}

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    /// Output format
    #[arg(long, default_value_t = Format::Object)]
    format: Format,

    /// Number of rows to skip
    #[arg(short, long, default_value_t = 0)]
    skip: usize,

    /// Filename
    #[arg(short, long)]
    filename: String,
}

fn get_object_keys(header: &[calamine::DataType]) -> Vec<String> {
    header
        .iter()
        .map(|x| x.to_string())
        .map(|x| x.replace(' ', "_").replace('.', ""))
        .map(|x| x.to_lowercase())
        .collect()
}

fn convert_values(value: &calamine::DataType) -> serde_json::Value {
    use serde_json::Value;
    match value {
        calamine::DataType::Int(x) => Value::Number(serde_json::Number::from(*x)),
        calamine::DataType::Float(x) => Value::Number(serde_json::Number::from_f64(*x).unwrap()),
        calamine::DataType::String(x) => Value::String(x.clone()),
        calamine::DataType::Bool(x) => Value::Bool(*x),
        calamine::DataType::Empty => Value::Null,
        _ => Value::String(value.to_string()),
    }
}

fn main() {
    let args = Args::parse();
    let mut file = calamine::open_workbook_auto(&args.filename).expect("Cannot open file");
    let sheet = &file.worksheets()[0].1;
    let mut result: Vec<serde_json::Value> = vec![];
    match args.format {
        Format::Object => {
            let first_row = &sheet[args.skip];
            for row in sheet.rows().skip(1 + args.skip) {
                let obj = serde_json::Value::Object(
                    get_object_keys(first_row)
                        .into_iter()
                        .zip(row.iter().map(convert_values))
                        .collect(),
                );
                result.push(obj);
            }
        }
        Format::List => {
            result = sheet
                .rows()
                .skip(args.skip)
                .into_iter()
                .map(|x| x.iter().map(convert_values).collect::<Vec<_>>())
                .map(|x| serde_json::Value::Array(x))
                .collect();
        }
    }
    serde_json::to_writer(std::io::stdout(), &result).expect("Cannot write JSON");
}
