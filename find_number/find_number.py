import os
import xlrd
import sys
from tkinter import *
from process_input import *

input_file_entry = None
sheet_entry = None
number_entry = None
occurence_entry = None

def result_window(msg):
    window = Toplevel()
    window.geometry('300x100')
    newlabel = Label(window, text = msg)
    newlabel.pack()

    button = Button(window, text = 'Quit', command=sys.exit)
    button.pack(pady = 10)

def recursive_find(worksheet, number, occurence):
    if 'output.txt' in os.listdir():
        with open('output.txt', 'w') as f:
            pass

    try:
        row = 0
        length = len(number)
        while True:
            index = worksheet.cell_value(row, 0)
            try:
                int(index)
            except Exception:
                print('跳過資料: {0}, 因為第{1}列/第{2}行不是數字'.format(index, row + 1, 1))
                row += 1
                continue

            count = 0
            for column in range(1,6):
                e = worksheet.cell_value(row, column)
                if e == '':
                    print('跳過期別{0}, 因為第{1}列/第{2}行, 欄位無資料'.format(int(index), row, column))
                    break

                if int(e) in number:
                    count += 1

            if (occurence is not None and count == occurence) or \
                    (occurence is None and count == length):
                with open('output.txt', 'a') as f:
                    f.write('第{0}期別中了 {1} 個數字\n'.format(int(index), count))

            row += 1
    except IndexError as e:
        print('結束. 計算結果放在 output.txt')
        result_window('結束. 計算結果放在 output.txt')


def callback():
    input_file = input_file_entry.get()
    sheet = sheet_entry.get()
    number = number_entry.get()
    occurence = occurence_entry.get()

    try:
        check_input_file(input_file)
        check_sheet(sheet)
        parsed_number = parse_number(number)
    except ValueError as e:
        print(e)
        result_window(e)
        return

    occurence = parse_occurence(occurence)

    workbook = xlrd.open_workbook(input_file)

    worksheet = workbook.sheet_by_index(int(sheet) - 1)

    recursive_find(worksheet, parsed_number, occurence)

if __name__ == "__main__":
   root = Tk()
   root.geometry("400x200")

   frame = Frame(root)
   frame.pack()

   input_file_entry = Entry(frame, width = 50)
   input_file_entry.insert(0,'輸入檔案名稱, 例如: abc.xlsx')
   input_file_entry.pack(padx = 5, pady = 5)

   sheet_entry = Entry(frame, width = 50)
   sheet_entry.insert(0,'要分析哪一個分頁, 例如: 1, 2, 3, 4...')
   sheet_entry.pack(padx = 5, pady = 5)

   number_entry = Entry(frame, width = 50)
   number_entry.insert(0,'要找哪些數字, 例如: 1, 4, 32')
   number_entry.pack(padx = 5, pady = 5)

   occurence_entry = Entry(frame, width = 50)
   occurence_entry.insert(0,'要指定中幾個數字, 例如: 1, 2, 3. 無法辨識成數字會全部列出')
   occurence_entry.pack(padx = 5, pady = 5)

   button = Button(frame, text = "calculate", command = callback)
   button.pack()

   root.mainloop()

