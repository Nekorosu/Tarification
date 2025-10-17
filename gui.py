import pandas as pd
import openpyxl


def main():
    # поиск колонки с преподавателями
    teacher_column = search_teacher_coll_on_file()


def search_teacher_coll_on_file(
    file_path="/home/geguh/VSCode/Tarification/xlsx/Нагрузка ПД.xlsx",
):

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
                return master_cell  # возвращаем координаты найденных ячеек
        else:
            master_cell = None

    except SyntaxError:
        if master_cell == None:
            print(
                "Ошибка. Не найдена соответствующая ячейка, проверьте верность введеных в ячейку данных"
            )
        else:
            print("Ошибка функции 'Поиск преподавателей в книге' ")


if __name__ == "__main__":
    main()
