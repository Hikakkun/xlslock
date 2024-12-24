use calamine::{open_workbook_from_rs, DataType, Reader, Xlsx};
use office_crypto::decrypt_from_file;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use umya_spreadsheet;
use xlslock::set_password;

/// テスト用の一時フォルダを作成
fn setup_test_folder() -> PathBuf {
    let test_folder = Path::new("./test_data");
    if !test_folder.exists() {
        fs::create_dir(test_folder).expect("Failed to create test folder");
    }
    test_folder.to_path_buf()
}

/// 新しいExcelファイルを作成し、指定されたセルに値を設定
fn create_test_excel(test_path: &Path, cell: (u32, u32), value: &str) {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book
        .get_sheet_by_name_mut("Sheet1")
        .expect("Failed to find 'Sheet1' in the new workbook");

    sheet.get_cell_mut(cell).set_value(value);

    umya_spreadsheet::writer::xlsx::write(&book, test_path)
        .expect("Failed to write new Excel file");
}

/// 指定された復号化済みデータを使用し、セルの値を検証
fn verify_excel(
    decrypted_data: Vec<u8>,
    expected_cell: (u32, u32),
    expected_value: &str
) {
    let cursor = Cursor::new(decrypted_data);
    let mut wb: Xlsx<_> = open_workbook_from_rs::<Xlsx<_>, _>(cursor)
        .expect("Failed to open decrypted workbook");
    let sheet = wb
        .worksheet_range("Sheet1")
        .expect("Failed to find 'Sheet1' in the decrypted workbook");
    let cell_value = sheet
        .get_value((expected_cell.0 -1, expected_cell.1 -1))
        .expect("Failed to get the value of the specified cell");

    
    assert_eq!(
        cell_value.get_string(),
        Some(expected_value),
        "The cell value does not match the expected value"
    );
}

#[test]
fn test_set_password_and_decrypt() {
    let test_folder = setup_test_folder();
    let test_path = test_folder.join("test.xlsx");
    let test_password = "test_password";
    let cell = (1, 1);
    // テスト用のExcelファイルを作成
    create_test_excel(&test_path, cell, "TEST1");

    // パスワードを設定
    set_password(&test_path, test_password)
        .expect("Failed to set password on the Excel file");

    //decrypt_from_file 自体はどんなパスワードを入れてもエラーにはならない
    // open_workbook_from_rs で正しく復号化されていなものをいれて初めてエラーになる
    let wrong_password_result = decrypt_from_file(&test_path, "xxyyzz").unwrap();
    assert!(open_workbook_from_rs::<Xlsx<_>, _>(Cursor::new(wrong_password_result)).is_err());

    // ファイルを復号化
    let decrypted_data  = decrypt_from_file(&test_path, test_password).expect("Failed to decrypt the file with the given password");

    // 復号化済みデータを検証
    verify_excel(decrypted_data , cell, "TEST1");
}
