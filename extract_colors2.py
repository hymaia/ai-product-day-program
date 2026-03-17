import openpyxl

wb = openpyxl.load_workbook('/Users/elsamargier/Desktop/AI product day/[AI Product Day 26] ProgrammeV2.xlsx')

# Color to category mapping (based on legend cells B24-B26)
# FFFFF2CC = yellow = "AI In Product" -> product
# FFCFE2F3 = blue = "AI for PMs & Designers" -> pm
# FFD9EAD3 = green = "Scaling AI Adoption" -> scaling
# FFD9D2E9 = purple = Keynote -> keynote
# FFEFEFEF = grey = Break/Intro -> skip
# FFF4CCCC = red/pink = exclusive workshop -> workshop

COLOR_MAP = {
    'FFFFF2CC': 'product',
    'FFCFE2F3': 'pm',
    'FFD9EAD3': 'scaling',
    'FFD9D2E9': 'keynote',
    'FFF4CCCC': 'workshop',
}

SKIP_COLORS = {'FFEFEFEF', 'FFF6F8F9', '00000000', 'FFFFFFFF', '00FFFFFF', 'FF000000'}

ws = wb['Programme']
print("=== PROGRAMME SHEET - Talk to Category Mapping ===\n")

results = []
for row in ws.iter_rows():
    for cell in row:
        fill = cell.fill
        if fill and fill.fgColor and fill.fgColor.type == 'rgb':
            color = fill.fgColor.rgb
            if cell.value and color not in SKIP_COLORS:
                val = str(cell.value).strip()
                # Skip legend entries and breaks
                if val in ('AI In Product', 'AI for PMs & Designers', 'Scaling AI Adoption', 'Break', 'Closing Words'):
                    continue
                category = COLOR_MAP.get(color, f'UNKNOWN({color})')
                results.append((cell.coordinate, color, category, val))

for coord, color, category, val in results:
    # Detect table rondes and ateliers
    val_lower = val.lower()
    if 'table ronde' in val_lower or 'table rond' in val_lower:
        category = 'ronde'
    elif 'atelier' in val_lower:
        category = 'workshop'
    print(f"[{coord}] {color} -> {category}")
    print(f"  {val[:100]}")
    print()

print("\n=== CLEAN MAPPING: title -> category ===\n")
for coord, color, category, val in results:
    val_lower = val.lower()
    if 'table ronde' in val_lower or 'table rond' in val_lower:
        category = 'ronde'
    elif 'atelier' in val_lower:
        category = 'workshop'
    # Clean title (remove flag emojis prefix if present)
    title = val
    print(f"{category:12} | {title[:90]}")
