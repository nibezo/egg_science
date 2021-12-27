from openpyxl import *

wb = load_workbook('EEG Science.xlsx')
eeg = wb.active
# eeg means the first worksheet
print(eeg['A1'].value)

for i in range(1, 100):
    print(eeg[f'A{i}'].value)
