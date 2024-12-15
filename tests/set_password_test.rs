extern crate umya_spreadsheet;
use office_crypto::decrypt_from_file;
use std::fs;
use std::path::Path;
use xlslock::set_password; // プロジェクト名をモジュールとして指定
#[test]
fn test_set_password() {
    let test_path = Path::new("test.xlsx");
    let test_password = "test_password";
    let mut book = umya_spreadsheet::new_file();
    book.get_sheet_by_name_mut("Sheet1")
        .unwrap()
        .get_cell_mut("A1")
        .set_value("TEST1");
    umya_spreadsheet::writer::xlsx::write(&book, &test_path).unwrap();

    // パスワード設定
    assert!(set_password(&test_path, test_password).is_ok());
    let result = decrypt_from_file(test_path, test_password);
    assert!(
        result.is_ok(),
        "Failed to decrypt the file with the given password"
    );
    fs::remove_file(&test_path).unwrap();
}
