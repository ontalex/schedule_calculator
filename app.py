from tkinter import Frame, Button, BOTH
from tkinter import ttk
from tkinter import filedialog

import pandas as pd
from pandastable import Table, TableModel


class TestApp(Frame):
    """Basic test frame for the table"""

    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master

        self.calculator = []
        self.group_list = []
        self.main.title("Калькулятор расписания")
        self.main.geometry("800x750+100+100")

        # панель страниц
        self.notebook = ttk.Notebook(self.main)
        self.frame_import = Frame(self.notebook)
        self.frame_analyze = Frame(self.notebook)
        self.notebook.add(self.frame_import, text="Import")
        self.notebook.add(self.frame_analyze, text="Analyze")

        # frame_import_controllers - фрейм с функциями
        self.frame_import_controllers = Frame(self.frame_import, borderwidth=1)
        self.frame_analyze_controllers = Frame(self.frame_analyze, borderwidth=1)

        # frame_import_table - фрейм с таблицей
        self.frame_import_table = Frame(self.frame_import, borderwidth=1)
        self.frame_analyze_table = Frame(self.frame_analyze, borderwidth=1)

        # Кнопка открытия новой таблицы
        btn_open_table = Button(
            self.frame_import_controllers,
            text="открыть таблицу",
            command=self.openTable,
        )
        # Кнопка очистки обрабатываемых данных
        btn_drop_table = Button(
            self.frame_import_controllers,
            text="Очистить таблицу",
            command=self.dropTable,
        )

        self.notebook.pack(fill=BOTH, expand=True)
        self.frame_import_controllers.pack(fill="x")
        self.frame_import_table.pack(fill="both", expand=True)
        self.frame_analyze_controllers.pack(fill="x")
        self.frame_analyze_table.pack(fill="both", expand=True)

        btn_open_table.pack(fill=BOTH, side="left", pady=2)
        btn_drop_table.pack(fill=BOTH, side="left", pady=2)

        # self.dataframe = None
        self.dataframe = pd.DataFrame(data=None)
        self.calc_dataframe = pd.DataFrame(columns=["Предмет", "Часы"], data=None)

        self.table = Table(
            self.frame_import_table,
            dataframe=self.dataframe,
            showstatusbar=True,
            showtoolbar=True,
            editable=False,
        )

        self.table_calc_0 = Table(
            self.frame_analyze,
            dataframe=self.calc_dataframe,
            showstatusbar=True,
            showtoolbar=True,
            editable=False,
        )

        return

    def selectGroup(self, event):
        select_index = 0

        # Поиск выбранной группы

        for group_index, group in enumerate(self.calculator):
            if group["current_group_name"] == self.select_group.get():
                select_index = group_index
                self.calc_dataframe = self.calculator[select_index]["group_lessons"]
                break

        self.table_calc_0.destroy()
        self.table_calc_0.update()
        self.table_calc_0 = tc0 = Table(
            self.frame_analyze_table,
            dataframe=pd.DataFrame.from_dict(
                self.calc_dataframe,
                "index",
                columns=["Предмет", "Часы"],
            ),
            showstatusbar=False,
        )
        tc0.show()

        self.table_calc_0.update()

    def getDataFrame(index_sheet=0):
        file_path = filedialog.askopenfilename()

        if file_path:
            dataframe = pd.read_excel(io=file_path, sheet_name=0)
            return dataframe

    def openTable(self):

        # Отчистка
        self.group_list = []
        for elem in self.frame_analyze_controllers.winfo_children():
            elem.destroy()

        template = {
            "current_group_name": None,
            "group_cords_x": 0,
            "group_cords_y": 0,
            "group_lessons": {},
        }

        self.dataframe = dataF = self.getDataFrame()

        # Приводим таблицу к нормальному виду
        self.tableEdit = TableModel(dataF)
        self.calculator = []

        # Начальные данные
        global pivot_x_cord
        global pivot_y_cord

        pivot_x_cord = 0
        pivot_y_cord = 0

        pivot_word = "Дни"

        max_y = self.tableEdit.getRowCount()
        max_x = self.tableEdit.getColumnCount()

        # Поиск опорной точки
        for y in range(0, int(max_y / 2)):
            for x in range(0, int(max_x / 2)):
                if str(self.tableEdit.getValueAt(x, y)) == pivot_word:
                    pivot_y_cord = x - 1
                    pivot_x_cord = y
                    break

        # определить группы и клонировать объекты
        for column_index in range(pivot_x_cord, max_x):
            cell = self.tableEdit.getValueAt(pivot_y_cord, column_index)

            if str(cell).find("-") != -1:
                # Нашли что-то похожее на группу
                new_element = template.copy()

                new_element["current_group_name"] = str(cell)
                new_element["group_cords_x"] = int(column_index)
                new_element["group_cords_y"] = int(pivot_y_cord)

                print("New Element: ", new_element, "\n\n")

                self.calculator.append(new_element)

        # вычислить калькуляцию дисциплин
        for group_index, group in enumerate(self.calculator):
            new_disct = {}

            for row_index in range(
                group["group_cords_y"] + 2, self.tableEdit.getRowCount()
            ):

                cell = self.tableEdit.getValueAt(
                    col=group["group_cords_x"], row=row_index
                ).strip()

                if cell != "":
                    if cell not in new_disct:
                        new_disct[cell] = [cell, 2]
                    else:
                        new_disct[cell][1] += 2

            self.calculator[group_index]["group_lessons"] = new_disct.copy()

        self.table = pt = Table(
            self.frame_import_table,
            dataframe=self.dataframe,
            showstatusbar=True,
            showtoolbar=True,
            editable=False,
        )
        pt.show()

        print("Data Calculators: ", self.calculator)

        # получить список групп
        for group_index, group in enumerate(self.calculator):
            self.group_list.append(group["current_group_name"])

        # Переустановка спиcка и таблицы
        self.select_group = ttk.Combobox(
            self.frame_analyze_controllers, values=self.group_list
        )
        self.select_group.pack(fill="x", expand=True)
        self.select_group.set(0)
        self.select_group.bind(
            "<<ComboboxSelected>>", lambda event: self.selectGroup(event)
        )

        self.table_calc_0.update()
        self.table.update()

    def dropTable(self):
        self.dataframe = None
        self.calc_dataframe = None

        self.table = pt = Table(
            self.frame_import_table,
            dataframe=self.dataframe,
            showstatusbar=True,
            showtoolbar=True,
            editable=False,
        )
        pt.show()

        self.table_calc_0 = tc0 = Table(
            self.frame_analyze_table,
            dataframe=self.calc_dataframe,
            showstatusbar=True,
            showtoolbar=False,
            editable=False,
        )
        tc0.show()

        self.group_list = []
        for elem in self.frame_analyze_controllers.winfo_children():
            elem.destroy()

        # Переустановка спиcка
        self.select_group = ttk.Combobox(
            self.frame_analyze_controllers, values=self.group_list
        )
        self.select_group.pack(fill="x", expand=True)
        self.select_group.set(0)
        self.select_group.bind(
            "<<ComboboxSelected>>", lambda event: self.selectGroup(event)
        )


app = TestApp()
# launch the app
app.mainloop()
