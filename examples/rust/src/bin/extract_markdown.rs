use std::{env, fs, path::Path};

fn detect_format(path: &Path) -> Option<&'static str> {
    let ext = path.extension()?.to_str()?.to_ascii_lowercase();
    match ext.as_str() {
        "docx" => Some("docx"),
        "xlsx" => Some("xlsx"),
        "pptx" => Some("pptx"),
        _ => None,
    }
}

fn main() {
    let mut args = env::args().skip(1);
    let Some(path_arg) = args.next() else {
        eprintln!("usage: extract_markdown <path-to-officemd>");
        std::process::exit(2);
    };

    let path = Path::new(&path_arg);
    let Some(format) = detect_format(path) else {
        eprintln!("unsupported file extension: {}", path.display());
        std::process::exit(2);
    };

    let content = match fs::read(path) {
        Ok(bytes) => bytes,
        Err(err) => {
            eprintln!("failed to read {}: {err}", path.display());
            std::process::exit(1);
        }
    };

    let markdown: Result<String, String> = match format {
        "docx" => officemd_docx::markdown_from_bytes(&content).map_err(|e| e.to_string()),
        "xlsx" => officemd_xlsx::markdown_from_bytes(&content).map_err(|e| e.to_string()),
        "pptx" => officemd_pptx::markdown_from_bytes(&content).map_err(|e| e.to_string()),
        _ => unreachable!(),
    };

    match markdown {
        Ok(md) => {
            print!("{md}");
        }
        Err(err) => {
            eprintln!("failed to render markdown for {}: {err}", path.display());
            std::process::exit(1);
        }
    }
}
