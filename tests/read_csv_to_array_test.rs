use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use xlslock::read_csv_to_array; // read_csv_to_array をインポート

#[test]
fn test_read_csv_to_array() {
    // 一時的なCSVファイルのパス
    let test_csv_path = Path::new("test.csv");

    // テスト用のCSVデータ
    let csv_data = "test1.xlsx,password1\ntest2.xlsx,password2\ninvalid_file.txt,password3\n";

    // 一時的なCSVファイルを作成
    let mut file = File::create(&test_csv_path).unwrap();
    file.write_all(csv_data.as_bytes()).unwrap();

    // 実行
    let result = read_csv_to_array(&test_csv_path);

    // 結果を検証
    assert!(result.is_ok(), "Failed to read CSV: {:?}", result);

    let data = result.unwrap();

    // データの内容を確認
    let expected: Vec<(PathBuf, String)> = vec![
        (PathBuf::from("test1.xlsx"), "password1".to_string()),
        (PathBuf::from("test2.xlsx"), "password2".to_string()),
    ];
    assert_eq!(data, expected, "Unexpected data read from CSV");

    // 一時的なCSVファイルを削除
    fs::remove_file(&test_csv_path).unwrap();
}
