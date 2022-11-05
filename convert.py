import openpyxl
from datetime import datetime


class parsing():
    def __init__(self):
        self.files = ["1_screen.xlsx", "2_screen.xlsx", "3_screen.xlsx", "4_screen.xlsx", "5_screen.xlsx",
                      "6_screen.xlsx", "7_screen.xlsx"]
        self.mask = []
        self.lessons = []
        self.day = []
        self.matrix = []
        self.time = [0, 1, 2, 3, 4, 5]
        self.data = {}
        self.groupList = []

    def save(self, sheet) -> None:
        self.matrix.clear()
        self.day.clear()
        self.lessons.clear()
        self.data.clear()
        self.mask.clear()

        for row in range(4, 41):
            days = sheet[row][0].value
            if type(days) == datetime:
                self.day.append(days.strftime("%d.%m.%Y"))

            self.mask.append(sheet[row][self.timetable].value)

        for row in self.mask:
            if row != None:
                self.lessons.append(row)
            else:
                row = 'Нет пары'
                self.lessons.append(row)

        try:
            while True:
                parsing.converting(self)
                self.matrix.append(self.matrixel)
        except IndexError:
            pass

        for day, urok in zip(self.day, self.time):
            self.data[day] = self.matrix[urok]

        parsing.write_in_txt(self)

    def write_in_txt(self) -> None:
        file_name = 'group_' + str(self.name_groups) + '.txt'
        self.groupList.append(str(self.name_groups))
        file = open(str('/home/sip/pars/parsing/Расписание/' + file_name), 'w+')

        for key in self.data:
            file.write(f'{key}, {self.data[key]}\n')
        file.close()

    def converting(self) -> None:
        self.matrixel = []
        self.matrixel.clear()
        while self.x < self.y:
            self.matrixel.append(self.lessons[self.x])
            self.x = self.x + 1
        else:
            self.y = self.y + 6
            return self.matrixel

    def writetimetable(self, file, x, y) -> None:
        book = openpyxl.open(file, read_only=True)
        sheet = book.active
        try:
            self.timetable = [2, 5, 8, 11, 14, 17, 20, 23]
            for self.name_groups, self.timetable in zip(self.name_groups, self.timetable):
                self.name_groups = sheet[2][self.name_groups].value
                self.x = x
                self.y = y
                parsing.save(self, sheet)
        except IndexError:
            pass

    def get_group_list(self) -> list:
        return self.groupList

    def main(self):
        for file in self.files:
            match int(file[0]):
                case 0 | 4:
                    x = 0
                    y = 7
                    self.name_groups = [0, 4, 7, 10, 13, 16, 19, 22]
                    parsing.writetimetable(self, file, x, y)
                case 5 | 6:
                    x = 0
                    y = 6
                    self.name_groups = [0, 5, 7, 10, 13, 16, 19, 22]
                    parsing.writetimetable(self, file, x, y)
                case _:
                    x = 0
                    y = 6
                    self.name_groups = [0, 4, 7, 10, 13, 16, 19, 22]
                    parsing.writetimetable(self, file, x, y)


if __name__ == "__main__":
    start_time = datetime.now()
    pars = parsing()
    pars.main()
    print(pars.get_group_list())
    print(datetime.now() - start_time)
