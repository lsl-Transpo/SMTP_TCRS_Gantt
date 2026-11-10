with open('SMTP_TCRS_Interactive_Gantt.html', encoding='utf-8') as f:
    text = f.read()

import re
# Find exportToExcel or exportToCSV
idx1 = text.find('function exportToExcel')
if idx1 != -1:
    print('=== EXPORT TO EXCEL ===')
    print(text[idx1:idx1+1500])

idx2 = text.find('function exportToCSV')
if idx2 != -1:
    print('=== EXPORT TO CSV ===')
    print(text[idx2:idx2+1000])

