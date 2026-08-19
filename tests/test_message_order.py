"""Тести для модуля генерації повідомлень (message_order)."""

from nodeautomationtoolkit.builtin_nodes.message_order import generate_decision_order


def test_generate_decision_order_basic():
    mapping = {
        "72 окрема механізована бригада": {
            "cipher": "А2167",
            "open_name": "72 окрема механізована бригада",
            "corps": "",
        },
    }
    text = (
        "НАКАЗ\n\n"
        "1. Капітана призначити до 72 окремої механізованої бригади.\n"
        "2. Солдата звільнити.\n\n"
        "Командир військової частини А0000"
    )
    res = generate_decision_order(text=text, mapping=mapping)
    decision = res["decision_text"]
    assert "А2167" in decision
    assert res["replaced_count"] >= 1
    # Check blank line rules: 1 before item, 2 before signer
    assert "\n\nКомандир" in decision or "\n\n\nКомандир" in decision


def test_generate_decision_order_custom_rules():
    text = "1. Старшого сержанта направити до ВЧ 1234."
    rules = "ВЧ 1234 -> військової частини А9999"
    res = generate_decision_order(text=text, rules=rules)
    assert "військової частини А9999" in res["decision_text"]
