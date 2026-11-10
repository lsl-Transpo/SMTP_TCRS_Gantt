with open('create_excel_schedule.py', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(30, 70):
    print(f"{i+1}: {lines[i]}", end='')

