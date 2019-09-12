import pandas as pd
import xlrd


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


table = Table("C:/Users/M3chanix/Desktop/НИИ Онкологии/ЛИС/12-08-2019_13-15-34/табл 2.xlsx")
table6 = Table6("C:/Users/M3chanix/Desktop/НИИ Онкологии/ЛИС/12-08-2019_13-15-34/табл 6.xlsx")
table3 = Table3("C:/Users/M3chanix/Desktop/НИИ Онкологии/ЛИС/12-08-2019_13-15-34/табл 3.xlsx")
table5 = Table3("C:/Users/M3chanix/Desktop/НИИ Онкологии/ЛИС/12-08-2019_13-15-34/табл 5.xlsx")
table7 = Table7("C:/Users/M3chanix/Desktop/НИИ Онкологии/ЛИС/12-08-2019_13-15-34/табл 7.xlsx")
print(table3.dataframe)



class Processor(object):
    def __init__(self, file_list):
        #todo: сделать правильную инициализацию для всех датафреймов
        self.file_list = file_list
        self.date = 0

    def get_date(self):
        with xlrd.open_workbook(self.file_list[0]) as wb:
            sheet = wb.sheet_by_index(0)
            self.date = sheet.cell_value(0, 1)

    # def process_files(self):
#todo: чтобы обрабатывать все файлы в одном цикле, нужно их различать и присваивать именно в свою именованую
#todo: переменную каждый. легче всего это сделать через уникальные имена для каждого файла