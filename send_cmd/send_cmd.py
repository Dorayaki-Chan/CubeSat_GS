import serial
import time
from tkinter import filedialog
import csv



def index(csvlsit):
    index = input('コマンド検索用語:')

    kouho = []
    count = 0
    print('番号','コマンド','説明')
    for retsu in csvList:
        if index in retsu[4]:
            kouho.append({'cmd':retsu[3], 'setsu':retsu[4]})
            print(count, kouho[count]['cmd'], kouho[count]['setsu'])

            count += 1
    return kouho

typ = [('csvファイル','*.csv;')] 
dir = '.\data'
fle = filedialog.askopenfilename(filetypes = typ, initialdir = dir)
with open(fle) as f:
    reader = csv.reader(f)
    count = 0
    csvList = []
    for row in reader:
        count += 1
        if count < 3:
            continue
        csvList.append(row)
        
# 2行目から
# csvList[2]
while True:
    kouho = index(csvList)
    if len(kouho)!=0:
        break
    print("何も見つかりませんでした")

number = input('送信したいコマンド番号を入力:')

# f f
# 1111 1111

# 16進数,
# 14byte
cmd = [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
cmd[3] = int(kouho[int(number)]['cmd'].replace('0x', ''), 16)
ser = serial.Serial("COM3", 9600)
print(ser.name)
print(cmd)
while True:
    x = input("送信しますか？Y/N")
    if(x=='Y'):
        ser.write(cmd)
    else:
        break
# time.sleep(0.2)
# result = ser.read_all()
# print(type(result))
# print(result)
ser.close()

