import openpyxl
from datetime import datetime
import time

class parsing():
    def __init__(self):
        self.files = ["1_screen.xlsx", "2_screen.xlsx", "3_screen.xlsx", "4_screen.xlsx", "5_screen.xlsx", "6_screen.xlsx", "7_screen.xlsx"]
        self.mask = []
        self.lessons = []
        self.day = []
        self.matrix = []
        self.time = [0,1,2,3,4,5]
        self.data = {}

    def save(self,sheet):
        self.matrix.clear()
        self.day.clear()
        self.lessons.clear()
        self.data.clear()
        self.mask.clear()


        for row in range(4, 41):
            days = sheet[row][0].value
            if type(days) == datetime:
                self.day.append(days.strftime("%d.%m.%Y"))

        for row in range(4, 41):
            lesson = sheet[row][self.timetable].value
            self.mask.append(lesson)

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

    def write_in_txt(self):
        file_name = 'group_'+str(self.name_groups)+'.txt'
        file = open(str('/home/user/parsing/Расписание/'+file_name), 'w+')
        for key in self.data:
            file.write(f'{key}, {self.data[key]}\n')
        file.close()
        
        

    def converting(self):
        self.matrixel = []
        self.matrixel.clear()
        while self.x < self.y:
            self.matrixel.append(self.lessons[self.x])
            self.x = self.x + 1
        else:
            self.y = self.y + 6
            return self.matrixel

    def writetimetable(self, file, x, y):
        book = openpyxl.open(file, read_only = True)
        sheet = book.active
        try:
            for self.name_groups, self.timetable in zip(self.name_groups, self.timetable):
                self.name_groups = sheet[2][self.name_groups].value
                self.x = x
                self.y = y
                parsing.save(self, sheet)
        except IndexError:
            pass
    
    def main(self):
        for file in self.files:
            if int(file[0]) in range(0, 5):
                x = 0
                y = 7
                self.name_groups = [0, 4, 7, 10, 13, 16, 19, 22]
                self.timetable = [2, 5, 8, 11, 14, 17, 20, 23]
                parsing.writetimetable(self, file, x, y)
            elif int(file[0]) in range (5,7):
                x = 0
                y = 6
                self.name_groups = [0, 5, 7, 10, 13, 16, 19, 22]
                self.timetable = [2, 5, 8, 11, 14, 17, 20, 23]
                parsing.writetimetable(self, file, x, y)
            else:
                x = 0
                y = 6
                self.name_groups = [0, 4, 7, 10, 13, 16, 19, 22]
                self.timetable = [2, 5, 8, 11, 14, 17, 20, 23]
                parsing.writetimetable(self, file, x, y)

if __name__ == "__main__":
    start_time = datetime.now()
    pars = parsing()
    pars.main()
    print(datetime.now() - start_time)
