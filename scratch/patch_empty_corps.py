import re

file_path = r'src/nodeautomationtoolkit/builtin_nodes/recipient_mapping.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: _build_sender_key returning empty string
content = content.replace(
    '''        elif corps_abbr in corps_resolved_cipher and corps_resolved_cipher[corps_abbr] == "":
            return ""''',
    '''        elif corps_abbr in corps_resolved_cipher and corps_resolved_cipher[corps_abbr] == "":
            return corps_col'''
)

# Fix 2: TCK naming. Replace "ОТЦК та СП" with the original region name without forcing "ОТЦК" when it shouldn't.
# Wait, let's just make it " ТЦК та СП" instead of " ОТЦК та СП" for Priority 2 and 3.
content = content.replace(
    '''        for stem, obl_name in _UKRAINE_OBLAST_STEMS.items():
            if low.startswith(stem):
                return f"{obl_name} ОТЦК та СП"
        # Перевіряємо чи це район відомої області
        for r_stem, obl_name in _RAYON_TO_OBLAST_MAP.items():
            if low.startswith(r_stem):
                return f"{obl_name} ОТЦК та СП"
        return f"{reg_nom} ОТЦК та СП"''',
    '''        # Якщо це справді область
        for stem, obl_name in _UKRAINE_OBLAST_STEMS.items():
            if low.startswith(stem) and ("облас" in text.lower() or "отцк" in text.lower()):
                return f"{obl_name} ОТЦК та СП"
        return f"{reg_nom} ТЦК та СП"'''
)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)
print("Patched.")
