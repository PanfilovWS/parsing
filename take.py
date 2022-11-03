print('Введите название группы')
file_name = 'group_'+input(str())+'.txt'
with open(str('/home/user/parsing/Расписание/'+file_name), 'r') as row:
    text = row.readlines()
print(text[0])