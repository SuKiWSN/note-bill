#coding:utf-8
import openpyxl
import time
from openpyxl.drawing.image import Image
import os
import PIL


if not os.path.exists('./清河'):
    os.mkdir('./清河')

fl = os.listdir('./清河/')

for fname in fl:
    deals = os.listdir('./清河/{}'.format(fname))
    for deal in deals:
        if deal.endswith('.txt'):
            continue
        wb = openpyxl.load_workbook('./清河/{}/{}'.format(fname, deal))
        sheet = wb.active
        row = 2
        all = 0
        while sheet['a{}'.format(str(row))].value != None:
            sum = float(sheet['d{}'.format(row)].value) * float(sheet['e{}'.format(row)].value)
            sheet['f{}'.format(row)] = '%.2f' % sum
            all += sum
            row += 1
        sheet['e{}'.format(row)] = '总计'
        sheet['f{}'.format(row)] = '%.2f' % all
        wb.save('./清河/{}/{}'.format(fname, deal))
        wb.close()


c_width = 15
r_height = 150

t = time.strftime('%Y.%m.%d',  time.localtime())
y, m, d = t.split('.')
if not os.path.exists('./清河/{}年'.format(y)):
    os.mkdir('./清河/{}年'.format(y))
if not os.path.exists('./清河/{}年/{}月.xlsx'.format(y, m)):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['a1'] = '日期'
    sheet['b1'] = '图片'
    sheet['c1'] = '款号'
    sheet['d1'] = '单价'
    sheet['e1'] = '件数'
    sheet['f1'] = '总价'
    sheet['g1'] = '款式'
    sheet['h1'] = '备注'
else:
    wb = openpyxl.load_workbook('./清河/{}年/{}月.xlsx'.format(y, m))
    sheet = wb.active

sheet.column_dimensions['B'].width = c_width
sheet.column_dimensions['A'].width = c_width
sheet.column_dimensions['G'].width = 30
all = 0
row = 2
while sheet['a{}'.format(row)].value != None:
    all += float(sheet['d{}'.format(row)].value) * float(sheet['e{}'.format(row)].value)
    sheet['f{}'.format(row)] = '%.2f' % (float(sheet['d{}'.format(row)].value) * float(sheet['e{}'.format(row)].value))
    row += 1

while 1:
    img = input('图片，输入 q 退出\n>')[1:-1]
    if img == '':
        break
    while not os.path.exists(img):
        img = input('图片不存在，请重新输入\n>')[1:-1]
    kuanhao = input('款号\n>')
    price = float(input('单价\n>'))
    price = '%.2f' % (price)
    num = float(input('数量\n>'))
    num = '%.2f' % (num)
    sum = '%.2f' % (float(price) * float(num))
    model = input('款式\n>')
    ad = input('备注\n')

    img = Image(img)
    sheet.row_dimensions[row].height = r_height
    img.width, img.height = (100, 200)
    sheet.add_image(img, 'b{}'.format(row))
    sheet['a{}'.format(row)] = t
    sheet['c{}'.format(row)] = kuanhao
    sheet['d{}'.format(row)] = price
    sheet['e{}'.format(row)] = num
    sheet['f{}'.format(row)] = sum
    sheet['g{}'.format(row)] = model
    sheet['h{}'.format(row)] = ad
    all += float(sum)
    row += 1
sheet['e{}'.format(row)] = '总计'
sheet['f{}'.format(row)] = '%.2f' % all
print('{}月总计{} Ԫ'.format(m, all))
wb.save('./清河/{}年/{}月.xlsx'.format(y, m))

fl = os.listdir('./清河/{}年'.format(y))
yearall = 0
for fname in fl:
    if fname.endswith('xlsx'):
        workspace = openpyxl.load_workbook('./清河/{}年/{}'.format(y, fname))
        sheet = workspace.active
        j = 2
        while sheet['e{}'.format(j)].value != None:
            j += 1
        yearall += float(sheet['f{}'.format(j-1)].value)
f = open('./清河/{}年/全年总计.txt'.format(y), 'w')
f.write('{} ￥'.format(yearall))
print('全年总计{} 元'.format(yearall))
input('按下回车退出程序\n')

f.close()
