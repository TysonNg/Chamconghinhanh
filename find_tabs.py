with open('templates/index.html', 'r', encoding='utf-8') as f:
    for i, line in enumerate(f):
        if 'tab-content' in line or 'id="tab-' in line or 'id=\'tab-' in line:
            print(f"{i+1}: {line.strip()}")
