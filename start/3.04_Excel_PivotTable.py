#Open Excel & Read value
from openpyxl import Workbook, load_workbook

wb = load_workbook('3.04_Excel_PivotTable.xlsx')
ws1 = wb['RAW']
ws2 = wb['Summary']

print(ws2['a3'].value + '|' + ws2['b3'].value)
print(ws2['a4'].value + '|' + str(ws2['b4'].value))
print(ws2['a5'].value + '|' + str(ws2['b5'].value))
print(ws2['a6'].value + '|' + str(ws2['b6'].value))

pivot = ws2._pivots[0] # any will do as they share the same cache
pivot.cache.refreshOnLoad = True

print(ws2['a3'].value + '|' + ws2['b3'].value)
print(ws2['a4'].value + '|' + str(ws2['b4'].value))
print(ws2['a5'].value + '|' + str(ws2['b5'].value))
print(ws2['a6'].value + '|' + str(ws2['b6'].value))

