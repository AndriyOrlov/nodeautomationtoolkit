import sys
import os
sys.path.insert(0, os.path.abspath('src'))
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _extract_tck_sender

tests = [
    "призваного Дніпровським районним у місті Києві територіальним центром комплектування та соціальної підтримки",
    "Шевченківським районним ТЦК та СП міста Києва",
    "Червоноградським РТЦК та СП Львівської області",
    "Солом'янським районним територіальним центром комплектування та соціальної підтримки",
]

for t in tests:
    print(f"Text: {t}")
    print(f"Result: {_extract_tck_sender(t)}")
    print()
