with open('SMTP_TCRS_Interactive_Gantt.html', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(5700, 5780):
    print(f"{i+1}: {lines[i]}", end='')

