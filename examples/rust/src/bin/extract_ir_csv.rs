use std::env;
use std::fs;
use std::io::Write;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).unwrap_or_else(|| {
        eprintln!(
            "Usage: cargo run -p officemd_examples --bin extract_ir_csv -- <path/to/file.csv>"
        );
        std::process::exit(2);
    });

    let content = fs::read(&path)?;
    let doc = officemd_csv::extract_tables_ir(&content)?;
    let stdout = std::io::stdout();
    let mut lock = stdout.lock();
    serde_json::to_writer_pretty(&mut lock, &doc)?;
    writeln!(&mut lock)?;

    Ok(())
}
