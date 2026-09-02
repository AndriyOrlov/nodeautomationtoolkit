import sys
sys.path.insert(0, r'src')
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _extract_tck_sender

text1 = "Тернопільського об’єднаного міського територіального центру комплектування та соціальної підтримки Тернопільської області"
t1 = _extract_tck_sender(text1)
print("text1 ->", t1.encode('unicode_escape').decode() if t1 else t1)

text2 = "Чортківського районного територіального центру комплектування та соціальної підтримки Тернопільської області"
t2 = _extract_tck_sender(text2)
print("text2 ->", t2.encode('unicode_escape').decode() if t2 else t2)
