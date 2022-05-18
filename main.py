import csv
import re

import xlsxwriter

workbook = xlsxwriter.Workbook('out.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True, 'font_color': 'red'})
bold.set_align('center')

bold_1 = workbook.add_format({'bold': True, 'font_color': 'black'})
bold_1.set_align('center')

bold_2 = workbook.add_format({'bold': True, 'font_color': 'blue'})
bold_2.set_align('center')

bold_3 = workbook.add_format({'bold': True, 'font_color': 'black'})
bold_3 = workbook.add_format({'bg_color': '#b4b4b4'})
bold_3.set_align('center')

data_format1 = workbook.add_format({'bg_color': '#b4b4b4'})
data_format1.set_align('center')
#
# =========================================================

# Format the first column
worksheet.set_column('A:A', 90, data_format1)
worksheet.set_column('B:B', 25)
worksheet.set_column('C:C', 25)

worksheet.set_default_row(25)

worksheet.write('A1', 'text', bold_3)
worksheet.write('B1', 'data_1', bold_1)
worksheet.write('C1', 'data_2', bold_1)



txt_ = (
    'RM37003-BA-Brejolandia.txt;odo de 22/10/2012 a 26/10/2012.',
    'ESPECIAL_01_PAC-MS-Douradina.txt;odo de 01Set2008 a 19Set2008,',
    '30-MG-Casa_Grande.txt;odo de 13Out2009 a 31Dez2009,',
    '09-SP-Maua.txt;odo de 17 a 21 de maio de 2004 sendo utilizados em sua execução técnicas e procedimentos,',
    '14-SP-Florinia.txt;odo de 29 de novembro a 02 de dezembro de 2004 sendo utilizados em sua execução , '
)


"""
(0?[1-9]|[12][0-9]|3[01])[\/\-\.](0?[1-9]|1[012])[ \/\.\-]

===============================================================

(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])-(?:0?[1-9]|1[0-2])-(?:19[0-9][0-9]|20[01][0-9])(?!\d)
Подробности

(?<!\d) - сразу слева не должно быть цифры
(?:0?[1-9]|[12][0-9]|3[01]) - целые числа от 1 до 31 (перед числами от 1 до 9 может быть необязательный 0)
- - дефис
(?:0?[1-9]|1[0-2]) - целые числа от 1 до 12 (перед числами от 1 до 9 может быть необязательный 0)
- - дефис
(?:19[0-9][0-9]|20[01][0-9]) - целые числа от 1900 до 2019
(?!\d) - сразу справа не должно быть цифры

===============================================================
((1[0-2]|0?[1-9])/(3[01]|[12][0-9]|0?[1-9])/(?:[0-9]{2})?[0-9]{2})|((Jan(uary)?|Feb(ruary)?|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?|Dec(ember)?)\s+\d{1,2},\s+\d{4})
===============================================================


===============================================================



===============================================================



===============================================================



re.findall('', txt_[1]))

"""



# https://regex101.com/r/YJvt3n/2

# print(re.findall(r'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])[\/\-\.](?:0?[1-9]|[12][0-9]|3[01])[\/\-\.](?:19[0-9][0-9]|20[01][0-9])(?!\d)', txt_[0]))
# print(re.findall(r'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])[a-zA-Z]{3}(?:19[0-9][0-9]|20[01][0-9])', txt_[3]))


# a = re.findall('\d+[,.]\d+', txt_[3])
#
# if not a:
#     print(f"List is empty: {a}")


# получим объект файла
with open("test.txt", "r") as file1:
    row = 2
    # итерация по строкам
    for line in file1:
        print(line.strip())
        #'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])[\/\-\.](?:0?[1-9]|[12][0-9]|3[01])([\/\-\.](?:19[0-9][0-9]|20[01][0-9]))?(?!\d)|(a \d\d){1}(?![a-zA-Z\/])'
        #'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])[\/\-\.](?:0?[1-9]|[12][0-9]|3[01])[\/\-\.]?(?:19[0-9][0-9]|20[01][0-9])?(?!\d)'
        a = re.findall(r'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])[\/\-\.](?:0?[1-9]|[12][0-9]|3[01])[\/\-\.]?(?:19[0-9][0-9]|20[01][0-9])?(?!\d)|(a \d\d){1}(?![a-zA-Z\/])', line)
        if not a:
            a = re.findall(r'(?<!\d)(?:0?[1-9]|[12][0-9]|3[01])[a-zA-Z]{3}(?:19[0-9][0-9]|20[01][0-9])', line)
            if not a:
                a = ('NONE', 'NONE')
        print(a)
        worksheet.write(f'A{row}', line)
        worksheet.write(f'B{row}', a[0])
        worksheet.write(f'C{row}', a[1])
        row += 1
workbook.close()

# row = 2
# for aaa in range(0, 3):
#
#     worksheet.insert_image(f'A{row}', file_name, {'x_scale': 0.25, 'y_scale': 0.25, 'x_offset': 10})
#     worksheet.write(f'B{row}', collection, bold_2) # worksheet.write_url(f'F{row}', url, string=f'{lot_num}')
#     worksheet.write(f'C{row}', volume)
#     worksheet.write(f'D{row}', d24h_)
#     worksheet.write(f'E{row}', d7d)
#     worksheet.write(f'F{row}', floor_price)
#     worksheet.write(f'G{row}', num_owners)
#     worksheet.write(f'H{row}', items)
#
#     row = row + 1
#
# workbook.close()



