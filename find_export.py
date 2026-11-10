with open('SMTP_TCRS_Interactive_Gantt.html', encoding='utf-8') as f:
    text = f.read()

import re
for m in re.finditer(r'export|xlsx|csv', text, re.IGNORECASE):
    idx = m.start()
    snippet = text[max(0, idx-40):min(len(text), idx+120)]
    if 'function' in snippet or 'addEventListener' in snippet or 'download' in snippet or 'XLSX.utils' in snippet:
        print(f"Line approx {text[:idx].count(chr(10))+1}: {snippet}")

