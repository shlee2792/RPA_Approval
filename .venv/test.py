import datetime
n = datetime.datetime.now()
n.isocalendar()
# (2017, 45, 1)
n = datetime.datetime(2022, 9, 1)
n.isocalendar()
# (2017, 44, 7)

print(n.isocalendar())