from nodeautomationtoolkit.builtin_nodes.output import format_result
from nodeautomationtoolkit.core.word_types import WordSaveResult


def test_formats_word_result_for_preview():
    result = WordSaveResult("C:/result.docx", 4)

    text = format_result(result)

    assert "C:/result.docx" in text
    assert "paragraph_count" in text
