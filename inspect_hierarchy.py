import json

with open('schedule_data.json', encoding='utf-8') as f:
    tasks = json.load(f)

print(f"Total tasks: {len(tasks)}")
for t in tasks[:25]:
    wbs = t.get('WBS', '')
    lvl = t.get('Level', '')
    pid = t.get('Parent ID', '')
    pt = t.get('Parent Task', '')
    name = t.get('Task Name', '')
    proj = t.get('Project', '')
    print(f"[{proj}] WBS: {wbs:<8} Lvl: {lvl} ParentID: {pid:<10} ParentTask: {pt:<30} Name: {name}")

levels = {}
for t in tasks:
    wbs = t.get('WBS', '')
    parts = len(wbs.split('.')) if wbs else 0
    levels[parts] = levels.get(parts, 0) + 1
print(f"\nWBS depth breakdown (dot count + 1): {levels}")

