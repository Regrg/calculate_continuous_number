from tkinter import *
import os
import xlrd
import sys

input_file_entry = None
sheet_entry = None
column_entry = None

def result_window(msg):
    window = Toplevel()
    window.geometry('300x100')
    newlabel = Label(window, text = msg)
    newlabel.pack()

    button = Button(window, text = 'Quit', command=sys.exit)
    button.pack(pady = 10)

def is_digit(data):
    try:
        int(data)
    except ValueError:
        return False

    return True

def count(worksheet, col):
    row = 1
    count_smaller_than_3 = 0
    try:
        while True:
            e = worksheet.cell_value(row, col)
            if (not is_digit(e)):
                with open('output_smaller_than_3.txt', 'a') as f:
                    f.write('結束. 最後還剩下 {0} 個小於3的數字\n\n'.format(count_smaller_than_3))
                return

            if (e >= 3):
                with open('output_smaller_than_3.txt', 'a') as f:
                    f.write('第 {0} 列是 {1}, 前面有 {2} 個小於3的數字\n'.format(row + 1, int(e), count_smaller_than_3))
                count_smaller_than_3 = 0
            else:
                count_smaller_than_3 += 1

            row += 1
    except IndexError as e:
        with open('output_smaller_than_3.txt', 'a') as f:
            f.write('結束. 最後還剩下 {0} 個小於3的數字\n\n'.format(count_smaller_than_3))

def recursive_count(worksheet, start_col):
    if 'output_smaller_than_3.txt' in os.listdir():
        with open('output_smaller_than_3.txt', 'w') as f:
            pass

    try:
        while True:
            e = worksheet.cell_value(0, start_col)
            if e == '':
                start_col += 1
                continue

            with open('output_smaller_than_3.txt', 'a') as f:
                try:
                    f.write('開頭第一行: {0}\n'.format(int(e)))
                except ValueError as e:
                    f.write('開頭第一行: {0}\n'.format(e))

            count(worksheet, start_col)
            start_col += 1
    except IndexError as e:
        print('結束. 計算結果放在 output_smaller_than_3.txt')
        result_window('結束. 計算結果放在 output_smaller_than_3.txt')

def callback():
    input_file = input_file_entry.get()
    sheet = sheet_entry.get()
    column = column_entry.get()

    if not input_file:
        print('輸入檔案不能空白')
        result_window('輸入檔案不能空白')
        return

    if input_file not in os.listdir():
        print('目錄下找不到檔案: {0}'.format(input_file))
        result_window('目錄下找不到檔案: {0}'.format(input_file))
        return

    if not sheet.isnumeric():
        print('分頁不是數字: {0}'.format(sheet))
        result_window('分頁不是數字: {0}'.format(sheet))
        return

    if int(sheet) < 1:
        print('分頁數字小於1: {0}'.format(sheet))
        result_window('分頁數字小於1: {0}'.format(sheet))
        return

    if not column.isnumeric():
        print('哪一行開始不是數字: {0}'.format(column))
        result_window('哪一行開始不是數字: {0}'.format(column))
        return

    if int(column) < 1:
        print('哪一行開始數字小於1: {0}'.format(column))
        result_window('哪一行開始小於1: {0}'.format(sheet))
        return

    workbook = xlrd.open_workbook(input_file)

    worksheet = workbook.sheet_by_index(int(sheet) - 1)

    recursive_count(worksheet, int(column) - 1)

if __name__ == "__main__":
   root = Tk()
   root.geometry("300x200")

   frame = Frame(root)
   frame.pack()

   label = Label(frame, text='計算數字小於3的間隔')
   label.pack(padx = 5, pady = 5)

   input_file_entry = Entry(frame, width = 50)
   input_file_entry.insert(0,'輸入檔案名稱, 例如: abc.xlsx')
   input_file_entry.pack(padx = 5, pady = 5)

   sheet_entry = Entry(frame, width = 50)
   sheet_entry.insert(0,'要分析哪一個分頁, 例如: 1, 2, 3, 4...')
   sheet_entry.pack(padx = 5, pady = 5)

   column_entry = Entry(frame, width = 50)
   column_entry.insert(0,'要從哪一行開始分析, 例如: 1, 10, 20...')
   column_entry.pack(padx = 5, pady = 5)

   button = Button(frame, text = "calculate", command = callback)
   button.pack()

   root.mainloop()

