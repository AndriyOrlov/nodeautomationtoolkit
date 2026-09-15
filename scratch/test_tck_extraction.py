import re
from src.nodeautomationtoolkit.builtin_nodes.recipient_mapping import _extract_tck_sender, _TCK_KEYWORDS_RE, _TCK_OBL_RE
texts = [
    'Ковельського районного територіального центру комплектування та соціальної підтримки Волинської області',
    'Івано-Франківського міського територіального центру комплектування та соціальної підтримки',
    'Вінницького обласного ТЦК та СП'
]
for t in texts:
    print(f'TEXT: {t}')
    print(f'  is_tck: {bool(_TCK_KEYWORDS_RE.search(t))}')
    print(f'  match obl: {_TCK_OBL_RE.search(t)}')
    print(f'  Result: {_extract_tck_sender(t)}\n')
