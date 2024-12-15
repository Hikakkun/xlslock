use clap::Parser;
use std::path::Path;
use xlslock::{collect_xlsx_paths, read_csv_to_array, set_password, get_password}; 
use env_logger;
use std::env;
use log::{error, info};

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    #[arg(help = "対象とするファイルもしくはディレクトリのパス")]
    path: String,
    #[arg(short = 'p', long, help= "コマンドライン引数としてパスワードを設定したい場合に使用(非推奨)")]
    password: Option<String>
}
fn process_xlsx(path: &Path, password: &str) {
    match set_password(path, password) {
        Ok(_) => info!("Processed file {:?} with password", path.display()),
        Err(e) => error!("Error processing {:?}: {:?}", path.display(), e),        
    }

}

fn process_csv(path : &Path){
    match read_csv_to_array(path) {
        Ok(data) => {
            for (file_path, password) in data {
                if file_path.exists() {
                    match set_password(&file_path, &password) {
                        Ok(_) => info!("Processed file {:?} with password", file_path.display()),
                        Err(e) => error!("Error processing {:?}: {:?}", file_path.display(), e),
                    }
                } else {
                    error!("File does not exist: {:?}", file_path.display());
                }
            }
        }
        Err(e) => {
            error!("Failed to process CSV: {:?}", e);
        }
    }
}

fn process_directory(path: &Path, password: &str) {
    for file_path in  collect_xlsx_paths(path){
        match set_password(&file_path, password) {
            Ok(_) => info!("Processed file {:?} with password", file_path.display()),
            Err(e) => error!("Error processing {:?}: {:?}", file_path.display(), e),
        }
    }
}

fn main() {
    env::set_var("RUST_LOG", "info");
    env_logger::init();
    let args = Args::parse();
    let path = Path::new(&args.path);
    let password = match args.password.as_deref() {
        Some(pwd) => pwd,
        None => {
            &match get_password(false) {
                Ok(pwd) => pwd,
                Err(err) => {
                    error!("Error: {}", err);
                    return;
                }
            }
        }
    };

    if path.is_dir(){
        process_directory(path, password);
        return;
    }

    if path.is_file() {
        match path.extension().and_then(|ext| ext.to_str()) {
            Some("xlsx") => {
                process_xlsx(path, password);
            }
            Some("csv") => {
                process_csv(path);
            }
            _ => {
                error!("Unsupported file type: {:?}", path);
            }
        }
        return;
    }

}