import sys
import os
sys.path.insert(0, os.path.abspath('src'))
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import map_military_units

mapping = {
    "22 окрема механізована бригада": {"cipher": "А2222", "abbreviation": "22 ОМБр"},
    "23 окрема механізована бригада": {"cipher": "А2323", "abbreviation": "23 ОМБр"},
    "Київський обласний ТЦК та СП": {"cipher": "Київський ОТЦК", "abbreviation": "Київський ОТЦК та СП"},
    "Шевченківський районний в місті Києві ТЦК та СП": {"cipher": "Шевченківський РТЦК", "abbreviation": "Шевченківський РТЦК"},
}

def test_case(name, text):
    print(f"--- {name} ---")
    res = map_military_units(text=text, mapping=mapping)
    for u, data in res["unit_paragraphs"].items():
        print(f"  Unit: {u}")
        for i in data["items"]:
            print(f"    Item: {i['text'][:50]}...")
    print()

test_case("1. Шапка ЗВІДКИ і перелік КУДИ", """НАКАЗ
§ 1
призначити із 22 окремої механізованої бригади до:
1. солдата Іванова, 23 окрема механізована бригада
""")

test_case("2. Шапка ЗВІДКИ до КУДИ", """НАКАЗ
§ 1
призначити із 22 окремої механізованої бригади до 23 окремої механізованої бригади:
1. солдата Петрова
""")

test_case("3. Шапка без ЗВІДКИ/КУДИ, все в списку", """НАКАЗ
§ 1
нижчезазначених:
1. солдата Сидорова, 22 окрема механізована бригада, призваного Шевченківським районним ТЦК та СП міста Києва
""")
