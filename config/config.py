# TODO: Результирующие наименования столбцов в выходном файле
RESULT_HEADERS = [
    "ФИО",
    "Должность",
    "Отдел",
    "Дата найма",
    "Зарплата"
]

LOWER_HEADERS = [column_name.lower() for column_name in RESULT_HEADERS]

RESULT_FILE_NAME = "Фильтр_столбца_{}.xlsx"

EXCEL_FORMAT = ["*.xlsx", "*.xls"]
