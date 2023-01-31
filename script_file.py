import os
import pandas
from openpyxl import load_workbook



class FilePath():
    # Формирование списка файлов из директории
    def generator(self):
        for root, dirs, files in list(os.walk(os.getcwd())):
            for file in files:
                if len(file.split('.')[0]) == 0:
                    file = ' ' + file
                yield os.path.join(root, file)

    def new_list(self, gen):
        # Список для дальнейшего использования, содержащий название папки, файла и расширение
        mylist = list()
        count = 1
        for i_item in gen:
            if not os.path.isdir(i_item):
                mylist.append((count, (os.path.dirname(i_item.split("\\")[0])), os.path.splitext(i_item)[0],
                               os.path.splitext(i_item.split("\\")[-1])[-1]))
            count += 1
        return mylist

    def Excel_file(self, data):
        # Создаем файл с полученными данными и сохраняем в Excel формате
        datafile = pandas.DataFrame(data, columns=['Номер строки', 'Папка файла', 'Название файла', 'Расширение файла'])
        try:
            with pandas.ExcelWriter('result.xlsx', engine='xlsxwriter') as file:
                datafile.to_excel(file, index=False)
        except Exception as ex:
            print(ex)


if __name__ == "__main__":
    test = FilePath()
    test.Excel_file(test.new_list(test.generator()))
