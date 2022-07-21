import calendar

e = 0
t = (calendar.weekday(2021, 7, 1) + 1)
q = 36
www = q
while q >= 0:
    w = 5 - t
    q -= w
    if q > 0:
        e += 2
        t = 0
print('Всего дней:', www + e)
print('Рабочих:', www)
print('Выходных:', e)
print(calendar.weekday(2021, 7, 1) + 1)