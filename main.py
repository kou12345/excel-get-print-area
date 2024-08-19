from src.excel_processor import process_excel_file


def main():
    # Excelファイルのパスを指定
    file_path = "/Users/koichiro-hira/Downloads/原稿・PDF/マルマンストア0703号/0703(水)号大均市祭・夏のカレーフェス.xlsx"

    # Excelファイルを処理
    results = process_excel_file(file_path)

    # 結果を出力
    for sheet_name, csv_data in results.items():
        print(f"--- Sheet: {sheet_name} ---")
        print(csv_data)


if __name__ == "__main__":
    main()
