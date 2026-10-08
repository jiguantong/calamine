// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use crate::datatype::{Data, DataRef, ExcelDateTime, ExcelDateTimeType};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
}

/// Check excel number format is datetime
pub fn detect_custom_number_format(format: &str) -> CellFormat {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;
    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => return CellFormat::Other, // first format only
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return CellFormat::DateTime,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return CellFormat::DateTime
            }
            _ => {
                if hms && s.eq_ignore_ascii_case(&prev) {
                    // ok ...
                } else {
                    hms = prev == '[' && matches!(s, 'm' | 'h' | 's' | 'M' | 'H' | 'S');
                }
            }
        }
        prev = s;
    }
    CellFormat::Other
}

pub fn builtin_format_by_id(id: &[u8]) -> CellFormat {
    // XLSX 的字符串 ID 与 XLS/XLSB 的数字 ID 共用分类，避免漏掉某一种读取路径。
    std::str::from_utf8(id)
        .ok()
        .and_then(|value| value.parse::<u16>().ok())
        .map_or(CellFormat::Other, builtin_format_by_code)
}

/// Check if code corresponds to builtin date format
pub fn builtin_format_by_code(code: u16) -> CellFormat {
    match code {
        // 27–36、50–58 是中日韩等地区的内置日期/时间格式，含中文完整日期 31。
        // 只依据格式 ID 分类，不根据数值大小猜日期；时长 46 继续单独处理。
        // 格式定义：https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat
        14..=22 | 27..=36 | 45 | 47 | 50..=58 => CellFormat::DateTime,
        46 => CellFormat::TimeDelta,
        _ => CellFormat::Other,
    }
}

// convert i64 to date, if format == Date
pub fn format_excel_i64(value: i64, format: Option<&CellFormat>, is_1904: bool) -> Data {
    match format {
        Some(CellFormat::DateTime) => Data::DateTime(ExcelDateTime::new(
            value as f64,
            ExcelDateTimeType::DateTime,
            is_1904,
        )),
        Some(CellFormat::TimeDelta) => Data::DateTime(ExcelDateTime::new(
            value as f64,
            ExcelDateTimeType::TimeDelta,
            is_1904,
        )),
        _ => Data::Int(value),
    }
}

// convert f64 to date, if format == Date
#[inline]
pub fn format_excel_f64_ref(
    value: f64,
    format: Option<&CellFormat>,
    is_1904: bool,
) -> DataRef<'static> {
    match format {
        Some(CellFormat::DateTime) => DataRef::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::DateTime,
            is_1904,
        )),
        Some(CellFormat::TimeDelta) => DataRef::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::TimeDelta,
            is_1904,
        )),
        _ => DataRef::Float(value),
    }
}

// convert f64 to date, if format == Date
pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> Data {
    format_excel_f64_ref(value, format, is_1904).into()
}

#[test]
fn builtin_locale_datetime_formats_are_recognized_by_both_readers() {
    for code in (27..=36).chain(50..=58) {
        assert_eq!(
            builtin_format_by_code(code),
            CellFormat::DateTime,
            "code={code}"
        );
        assert_eq!(
            builtin_format_by_id(code.to_string().as_bytes()),
            CellFormat::DateTime,
            "id={code}"
        );
    }
}

#[test]
fn builtin_numeric_and_duration_formats_keep_their_types() {
    for code in [0, 1, 2, 9, 10, 23, 26, 37, 44, 48, 49, 59, 164] {
        assert_eq!(
            builtin_format_by_code(code),
            CellFormat::Other,
            "code={code}"
        );
        assert_eq!(
            builtin_format_by_id(code.to_string().as_bytes()),
            CellFormat::Other,
            "id={code}"
        );
    }
    assert_eq!(builtin_format_by_code(46), CellFormat::TimeDelta);
    assert_eq!(builtin_format_by_id(b"46"), CellFormat::TimeDelta);
}

#[test]
fn builtin_locale_dates_respect_workbook_epoch_and_time() {
    let format = builtin_format_by_code(31);
    for (serial, is_1904) in [(46024.5, false), (44562.5, true)] {
        let Data::DateTime(value) = format_excel_f64(serial, Some(&format), is_1904) else {
            panic!("日期格式不能退化为普通数字");
        };
        assert_eq!(value.to_ymd_hms_milli(), (2026, 1, 2, 12, 0, 0, 0));
        let Data::DateTime(value) = format_excel_i64(serial as i64, Some(&format), is_1904) else {
            panic!("整数日期序列号也必须按日期读取");
        };
        assert_eq!(value.to_ymd_hms_milli(), (2026, 1, 2, 0, 0, 0, 0));
    }
}

/// Ported from openpyxl, MIT License
/// https://foss.heptapod.net/openpyxl/openpyxl/-/blob/a5e197c530aaa49814fd1d993dd776edcec35105/openpyxl/styles/tests/test_number_style.py
#[test]
fn test_is_date_format() {
    assert_eq!(
        detect_custom_number_format("DD/MM/YY"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("H:MM:SS;@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0\\ [$\\u20bd-46D]"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("m\"M\"d\"D\";@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm:ss"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("#,##0\\ [$''u20bd-46D]"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("0_ ;[Red]\\-0\\ "),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    assert_eq!(
        detect_custom_number_format("#,##0.0####\" YMD\""),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("[h]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[ss]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[s].000"),
        CellFormat::TimeDelta
    );
    assert_eq!(detect_custom_number_format("[m]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[mm]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[Blue]\\+[h]:mm;[Red]\\-[h]:mm;[Green][h]:mm"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta][s].00"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm;[=0]\\-"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta].00"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta]General"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("ha/p\\\\m"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0.00\\ _M\"H\"_);[Red]#,##0.00\\ _M\"S\"_)"),
        CellFormat::Other
    );
}
