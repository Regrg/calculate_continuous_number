from tkinter import *
import os
import xlrd
import sys

input_file_entry = None
sheet_entry = None
number_entry = None

def parse_number(number):
    result = number.replace(' ', '').split(",")
    for v in result:
        if not v.isnumeric():
            return None

    return [int(x) for x in result]

def result_window(msg):
    window = Toplevel()
    window.geometry('300x100')
    newlabel = Label(window, text = msg)
    newlabel.pack()

    button = Button(window, text = 'Quit', command=sys.exit)
    button.pack(pady = 10)

def recursive_match(worksheet, number):
    if 'output.txt' in os.listdir():
        with open('output.txt', 'w') as f:
            pass

    result_list = []
    try:
        row = 1
        while True:
            index = worksheet.cell_value(row, 0)
            length = len(number)
            count = 0
            for column in range(1,6):
                e = worksheet.cell_value(row, column)
                if e == '':
                    print('跳過期別{0}, 因為第{1}列/第{2}行, 欄位無資料'.format(int(index), row, column))
                    break

                if int(e) in number:
                    count += 1

            if count == length:
                result_list.append(int(index))

            row += 1
    except IndexError as e:
        pass

    length = len(result_list)
    if length == 0:
        pass
    elif length == 1:
        with open('output.txt', 'a') as f:
            f.write('出現在{0}期別\n'.format(result_list[0]))
    else:
        with open('output.txt', 'a') as f:
            for i in range(length - 1):
                f.write('出現在{0}期別, 跟下一期隔了 {1} 期\n'.format(result_list[i], result_list[i + 1] - result_list[i]))
                if i == length - 2:
                    f.write('出現在{0}期別\n'.format(result_list[i + 1]))

    print('結束. 計算結果放在 output.txt')
    result_window('結束. 計算結果放在 output.txt')

def callback():
    input_file = input_file_entry.get()
    sheet = sheet_entry.get()
    number = number_entry.get()

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

    parsed_number = parse_number(number)
    if not parsed_number:
        print ('要找的數字無法辨識: {0}. 數字間請用逗號隔開'.format(number))
        result_window('要找的數字無法辨識: {0}. 數字間請用逗號隔開'.format(number))
        return

    workbook = xlrd.open_workbook(input_file)

    worksheet = workbook.sheet_by_index(int(sheet) - 1)

    recursive_match(worksheet, parsed_number)

if __name__ == "__main__":
   root = Tk()
   root.geometry("300x200")

   frame = Frame(root)
   frame.pack()

   input_file_entry = Entry(frame, width = 50)
   input_file_entry.insert(0,'輸入檔案名稱, 例如: abc.xlsx')
   input_file_entry.pack(padx = 5, pady = 5)

   sheet_entry = Entry(frame, width = 50)
   sheet_entry.insert(0,'要分析哪一個分頁, 例如: 1, 2, 3, 4...')
   sheet_entry.pack(padx = 5, pady = 5)

   number_entry = Entry(frame, width = 50)
   number_entry.insert(0,'要找哪些數字(最多五個), 例如: 1, 4, 32')
   number_entry.pack(padx = 5, pady = 5)

   button = Button(frame, text = "calculate", command = callback)
   button.pack()

   root.mainloop()

