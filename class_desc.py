import pandas as pd
import xlrd
import os
import shutil
import numpy as np


class Table:
    """For tables 1 and 2"""

    def __init__(self, file):
        self.dataframe = pd.read_excel(file, header=None)

        # drop last column and row
        self.dataframe = self.dataframe.drop(len(self.dataframe) - 1)
        self.dataframe = self.dataframe.drop(columns=self.dataframe.columns[-1])
        self.process()

    def process(self):
        self.date = self.dataframe.iat[0, 2]
        self.dataframe.columns = self.dataframe.iloc[1]
        self.dataframe = self.dataframe.drop([0, 1, 2])
        self.dataframe = self.dataframe.reset_index(drop=True)
        self.dataframe = self.dataframe.drop(columns="Подразделения лаборатории")


class Table6(Table):
    def process(self):
        self.date = self.dataframe.iat[0, 3]
        self.dataframe.columns = self.dataframe.iloc[1]
        self.dataframe = self.dataframe.drop([0, 1, 2])
        self.dataframe = self.dataframe.reset_index(drop=True)
        self.dataframe = self.dataframe.drop(columns="Подразделения лаборатории")
        # сохраняем нахвание отделения, для которого делалась выписка
        self.dataframe.iat[0, 1] = self.dataframe.iat[0, 0]
        self.dataframe = self.dataframe.drop(columns="Отделения заказчика")


class Table3(Table):
    """For tables 3 and 5"""

    def process(self):
        self.date = self.dataframe.iat[0, 2]
        self.dataframe.iat[1, 2] = "Всего"
        self.dataframe.iat[4, 1] = self.dataframe.iat[4, 0]
        self.dataframe.iat[1, 1] = self.dataframe.iat[2, 1]
        self.dataframe.columns = self.dataframe.iloc[1]
        self.dataframe = self.dataframe.drop([0, 1, 2, 3])
        self.dataframe = self.dataframe.reset_index(drop=True)
        self.dataframe = self.dataframe.drop(columns=self.dataframe.columns[0])


class Table7(Table):
    def process(self):
        self.date = self.dataframe.iat[0, 3]
        self.dataframe.iat[3, 3] = self.dataframe.iat[1, 3]
        self.dataframe.iat[3, 4] = self.dataframe.iat[2, 4]
        self.dataframe.iat[3, 5] = self.dataframe.iat[2, 5]
        self.dataframe.iat[5, 2] = self.dataframe.iat[5, 1]
        self.dataframe.columns = self.dataframe.iloc[3]
        self.dataframe = self.dataframe.drop([0, 1, 2, 3, 4])
        self.dataframe = self.dataframe.reset_index(drop=True)
        self.dataframe = self.dataframe.drop(columns=[self.dataframe.columns[0], self.dataframe.columns[1]])


class Department:
    def __init__(self, path):
        #        path - путь папки, в которой лежат файлы 5,6 и 7 для этого отделения
        self.table5 = Table3(path + "/5.xlsx")
        self.table6 = Table6(path + "/6.xlsx")
        self.table7 = Table7(path + "/7.xlsx")
        self.process_6_7()
        self.process_3_5()

    def process_3_5(self):
        self.processed_3_5 = self.table5.dataframe.copy()
        for column in self.processed_3_5.columns[2:]:
            self.processed_3_5[column] = self.processed_3_5[column] / self.processed_3_5[self.processed_3_5.columns[1]] \
                                         * 100

    def process_6_7(self):
        self.processed_6_7 = pd.merge(self.table6.dataframe, self.table7.dataframe, on="Микроорганизмы")


class Processor(object):
    def __init__(self, path):
        # todo: сделать правильную инициализацию для всех датафреймов
        self.table1 = Table(path + "/1.xlsx")
        self.table2 = Table(path + "/2.xlsx")
        self.table3 = Table3(path + "/3.xlsx")

        self.department_dict = {}
        for i in range(1, self.table3.dataframe.shape[0]):
            self.department_dict[self.table3.dataframe[self.table3.dataframe.columns[0]].iloc[i]] = \
                "department_" + str(i)

        for department in self.department_dict:
            self.__dict__[self.department_dict[department]] = Department(path + "/" + str(department))

        self.process_1_2()
        self.process_2_3()
        self.process_2_4()

    def process_1_2(self):
        self.processed_1_2 = self.table1.dataframe.copy()
        self.processed_1_2[" "] = self.table2.dataframe["Количество"]
        self.processed_1_2["%"] = self.processed_1_2[" "] / self.processed_1_2["Количество"] * 100

    def process_2_3(self):
        self.processed_2_3 = self.table3.dataframe.copy()
        for column in self.processed_2_3.columns[2:]:
            self.processed_2_3[column] = self.processed_2_3[column] / self.processed_2_3[self.processed_2_3.columns[1]] \
                                         * 100
        self.processed_2_3 = self.processed_2_3

    def process_2_4(self):
        for i in range(1, self.table3.dataframe.shape[0]):
            # first converting series to dataframe
            series = self.processed_2_3.iloc[i]
            dataframe = series.to_frame()
            dataframe.insert(0, "index", series.index)
            self.__dict__[self.department_dict[self.processed_2_3 \
                [self.processed_2_3.columns[0]].iloc[i]]].__dict__["processed_2_4"] = dataframe
