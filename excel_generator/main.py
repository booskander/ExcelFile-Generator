import csv
import random
import xlsxwriter

class ExcelGenerator:
    def __init__(self, filename, numberOfItems, lower, upper):
        self.__workbook = xlsxwriter.Workbook(filename)
        self.__workSheet = self.__workbook.add_worksheet()
        self.row = 0
        self.col = 0
        self.N = numberOfItems
        self.lower = lower
        self.upper = upper

    def __generateRandom(self) -> int:
        return random.randint(self.lower, self.upper)

    def __fillUp(self):
        for i in range(self.N):
            self.__workSheet.write(self.row, self.col, self.__generateRandom())
            self.row += 1
        self.__workbook.close()

    def getSheet(self):
        self.__fillUp()
        return self.__workSheet


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    myGen = ExcelGenerator("Array2.xlsx", 10, 0, 100)
    myGen.getSheet()
