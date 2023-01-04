import os

def check_input_file(input_file):
    if not input_file:
        raise ValueError('輸入檔案不能空白')

    if input_file not in os.listdir():
        raise ValueError('目錄下找不到檔案: {0}'.format(input_file))

def check_sheet(sheet):
    if not sheet.isnumeric():
        raise ValueError('分頁不是數字: {0}'.format(sheet))

    if int(sheet) < 1:
        raise ValueError('分頁數字小於1: {0}'.format(sheet))

def parse_number(number):
    result = number.replace(' ', '').split(",")
    for v in result:
        if not v.isnumeric():
            raise ValueError('要找的數字無法辨識: {0}. 數字間請用逗號隔開'.format(number))

    result = [int(x) for x in result]
    if len(result) != len(set(result)):
        raise ValueError('要找的數字有重複: {0}'.format(number))

    return result

def parse_occurence(occurence):
    if not occurence.isnumeric():
        return None

    return int(occurence)
