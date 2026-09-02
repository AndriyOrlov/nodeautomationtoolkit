import sys
sys.path.insert(0, r'src')
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import _build_unit_fuzzy_pattern
import re

text = "мінометній батареї механізованого батальйону 111 окремої механізованої бригади"
pattern = _build_unit_fuzzy_pattern("111 окрема механізована бригада")
print("Pattern:", pattern.pattern)
match = pattern.search(text)
print("Match found:", match is not None)

text2 = "мінометній батареї механізованого батальйону ХХХ окремої механізованої бригади"
pattern2 = _build_unit_fuzzy_pattern("ХХХ окрема механізована бригада")
print("Pattern2:", pattern2.pattern)
match2 = pattern2.search(text2)
print("Match2 found:", match2 is not None)
