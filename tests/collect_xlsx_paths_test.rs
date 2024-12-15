use std::fs::{self, File};
use std::path::PathBuf;
use xlslock::collect_xlsx_paths; // collect_xlsx_paths をインポート

#[test]
fn test_collect_xlsx_paths() {
    // 一時的なディレクトリを作成
    let temp_dir = tempfile::tempdir().unwrap();
    let temp_dir_path = temp_dir.path();

    // テスト用ファイルの作成
    let valid_xlsx_1 = temp_dir_path.join("test1.xlsx");
    let valid_xlsx_2 = temp_dir_path.join("subdir/test2.xlsx");
    let invalid_txt = temp_dir_path.join("test.txt");
    let subdir = temp_dir_path.join("subdir");

    // サブディレクトリを作成
    fs::create_dir(&subdir).unwrap();

    // ファイルを作成
    File::create(&valid_xlsx_1).unwrap();
    File::create(&valid_xlsx_2).unwrap();
    File::create(&invalid_txt).unwrap();

    // collect_xlsx_paths を実行
    let result = collect_xlsx_paths(temp_dir_path);

    // 結果を検証
    let expected: Vec<PathBuf> = vec![
        valid_xlsx_1.canonicalize().unwrap(),
        valid_xlsx_2.canonicalize().unwrap(),
    ];

    assert_eq!(result, expected, "Unexpected result from collect_xlsx_paths");
}
