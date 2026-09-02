import sys
sys.path.insert(0, r'src')
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

text = '''
§ 2
1.1. солдата ІВАНОВА, призначити в 5 окрему штурмову бригаду.
'''

mapping = {
    "10 АК": {"open_name": "10 Армійський корпус", "abbreviation": "10 АК", "corps": ""},
    "5 ОШБр": {"open_name": "5 окрема штурмова бригада", "abbreviation": "5 ОШБр", "corps": "10 АК"}
}

res = map_military_units(text=text, mapping=mapping)
print(res)
