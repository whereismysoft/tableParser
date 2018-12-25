import xlrd
from collections import OrderedDict
import simplejson as json
import sys

wb = xlrd.open_workbook('dealers.xlsx', encoding_override="utf-8")
sh = wb.sheet_by_index(0)

arr = []
keys = tuple(('id', 'title', 'address', 'addressLink','phones', 'emails', 'location'))

for rownum in range(1, sh.nrows):
    values = []
    row_values = sh.row_values(rownum)
    id = rownum
    other = row_values[2].split(',')
    title = other[0]
    address = ' '.join(other[1:len(other)-1])
    addressLink = ''
    phones = []
    phones.append(other[len(other) - 1])
    emails = ''
    location = ''
    values.append(id)
    values.append(title)
    values.append(address)
    values.append(addressLink)
    values.append(phones)
    values.append(emails)
    values.append(location)
    arr.append(dict(zip(keys, values)))


j = json.dumps(arr, ensure_ascii=False, encoding='utf8')
print(j)

with open('data.json', 'w') as f:
    f.write(j)
