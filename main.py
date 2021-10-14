import sys
from openpyxl import load_workbook

filename = 'export.xlsx'

# Columns in Excel:
skus_excel = 'A'
categories_excel = 'E'
product_online_excel = 'H'

try:
    workbook = load_workbook(filename=filename)
    sheet = workbook.active
except FileNotFoundError:
    sys.exit(f'Er is geen bestand met de naam \'{filename}\' gevonden.\n')

# check values
if sheet[skus_excel + '1'].value != 'sku':
    sys.exit(f'Kolom {skus_excel} is niet sku.')
if sheet[categories_excel + '1'].value != 'categories':
    sys.exit(f'Kolom E {categories_excel} is niet categories excel.')
if sheet[product_online_excel + "1"].value != 'product_online':
    sys.exit(f'Kolom H {product_online_excel} is niet product_online.')

# iteration
skus = [skus_iter.value for skus_iter in sheet[skus_excel]]
categories = [categories_iter.value for categories_iter in sheet[categories_excel]]
product_online = [product_online_iter.value for product_online_iter in sheet[product_online_excel]]

while True:
    category = input('Uit welke categorie moeten de sku\'s komen? ')
    row = 0
    count = 0
    online_count = 0
    for sku in skus:
        if(categories[row] is not None) and (category in categories[row]):
            online = 'online' if product_online[row] == 1 else ''
            print(f'{skus[row]},{online}')
            if (product_online[row] == 1):
                online_count += 1
            count += 1
        row += 1
    if count <= 0:
        print('Geen producten gevonden. Denk aan hoofdletters.\n')
    elif count > 0:
        print(str(count) + (' product' if count == 1 else ' producten') \
              + f' gevonden. {online_count} hiervan online.\n')
