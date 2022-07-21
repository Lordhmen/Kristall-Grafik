import sqlite3

conn = sqlite3.connect("BDGrafik.db")
cursor = conn.cursor()
q = 117
for i in range(162, 170):
    q += 1
    h = (q, i)
    cursor.execute("INSERT INTO Autoclava VALUES (?, ?)", h)
conn.commit()
