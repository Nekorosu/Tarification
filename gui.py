import pandas as pd
import openpyxl


file_path = "/home/geguh/VSCode/Tarification/xlsx/Нагрузка ПД.xlsx"


def main():

    # поиск колонки с преподавателями
    teacher_column_coordinates = search_teacher_coll_on_file()

    print(extract_teachers_from_column(teacher_column_coordinates))


def search_teacher_coll_on_file():

    try:
        searched_cell = "преподаватель"

        wb = openpyxl.load_workbook(file_path)  # загружаем рабочую книгу excel
        ws = wb.active  # ставим рабочую страницу книги

        merged_cells = (
            ws.merged_cells.ranges
        )  # определяем все объединенные ячейки страницы

        for cell in merged_cells:  # перебираем все объединенные ячейки
            start_row = cell.min_row  # стартовая ячейка строки
            start_col = cell.min_col  # стартовая ячейка колонны

            master_cell = ws.cell(row=start_row, column=start_col).value.strip().lower()

            # задаем переменной значение мастер ячейки, состоящей из старт. строк и колонн

            if master_cell == searched_cell:
                # проверяем мастер ячейку на соответствие переменной, которую ищем

                coordinate_of_merged_cells = [start_col, start_row]
                return (
                    coordinate_of_merged_cells  # возвращаем координаты найденных ячеек
                )
        else:
            master_cell = None

    except SyntaxError:
        if master_cell == None:
            return "Ошибка. Не найдена соответствующая ячейка, проверьте верность введеных в ячейку данных."

        else:
            return "Ошибка функции 'Поиск преподавателей в книге'."


def extract_teachers_from_column(coordinates_of_teachers_column):

    df = pd.read_excel(file_path, header=None)
    column_num = coordinates_of_teachers_column[0] - 1
    row_num = coordinates_of_teachers_column[1] - 1
    teachers = df.iloc[row_num + 1 :, column_num].dropna().tolist()
    return teachers


if __name__ == "__main__":
    main()
