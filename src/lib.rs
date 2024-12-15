use log::error;
use rpassword::read_password;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::{Path, PathBuf};
use walkdir::WalkDir;

pub fn get_password(single_input: bool) -> Result<String, String> {
    if single_input {
        eprint!("Enter your password: ");
        match read_password() {
            Ok(password) => Ok(password),
            Err(err) => Err(format!("Failed to read password: {}", err)),
        }
    } else {
        eprint!("Enter your password: ");
        let password = match read_password() {
            Ok(pwd) => pwd,
            Err(err) => return Err(format!("Failed to read password: {}", err)),
        };

        eprint!("Confirm your password: ");
        let confirm_password = match read_password() {
            Ok(pwd) => pwd,
            Err(err) => return Err(format!("Failed to read password: {}", err)),
        };

        if password == confirm_password {
            Ok(password)
        } else {
            Err("Passwords do not match.".to_string())
        }
    }
}

pub fn set_password(
    path: &std::path::Path,
    password: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    // ファイルを読み込む
    let book = umya_spreadsheet::reader::xlsx::read(path)?;

    // パスワードを設定して保存
    umya_spreadsheet::writer::xlsx::write_with_password(&book, path, password)?;

    Ok(())
}

pub fn collect_xlsx_paths(dir: &Path) -> Vec<PathBuf> {
    WalkDir::new(dir)
        .into_iter()
        .filter_map(|entry| entry.ok()) // WalkDirのエラーをスキップ
        .filter(|entry| entry.path().extension().map_or(false, |ext| ext == "xlsx")) // 拡張子がxlsxであるか確認
        .filter_map(|entry| entry.path().canonicalize().ok()) // 正規化されたパスを取得
        .collect()
}
type CsvData = Vec<(PathBuf, String)>;

// CSVを読み込んで配列に保存する関数
pub fn read_csv_to_array(path: &Path) -> Result<CsvData, Box<dyn std::error::Error>> {
    let file = File::open(path)?;
    let reader = BufReader::new(file);
    let mut data = Vec::new();

    for (i, line) in reader.lines().enumerate() {
        match line {
            Ok(content) => {
                let fields: Vec<&str> = content.split(',').collect();
                if fields.len() != 2 {
                    error!("Invalid format in CSV on line {}: {:?}", i + 1, content);
                    continue;
                }

                let file_path = PathBuf::from(fields[0]);
                let password = fields[1].to_string();

                if file_path.extension().map_or(false, |ext| ext == "xlsx") {
                    data.push((file_path, password));
                } else {
                    error!("Unsupported file type on line {}: {:?}", i + 1, file_path);
                }
            }
            Err(e) => error!("Error reading CSV line {}: {:?}", i + 1, e),
        }
    }

    Ok(data)
}
