data = ['tegrt', None, 'tgtrg']
data2 = []
for a in data:
    if a != None:
        data2.append(a)
    else:
        a = 'no'
        data2.append(a)

print(data2)