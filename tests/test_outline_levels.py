import unittest

from docx_io import _builtin_styles, docx_bytes_to_document, document_to_docx_bytes


class OutlineLevelRoundTripTests(unittest.TestCase):
    def test_builtin_styles_are_ordered_by_outline_level_without_code(self):
        styles = _builtin_styles()

        self.assertEqual(
            [style["id"] for style in styles],
            ["Normal", "Heading1", "Heading2", "Heading3", "NormalL1", "NormalL2", "NormalL3"],
        )

    def test_extended_outline_levels_survive_docx_round_trip(self):
        document = {
            "blocks": [
                {
                    "type": "paragraph",
                    "style_id": style_id,
                    "alignment": "align_left",
                    "runs": [{"text": style_id, "descriptor": ["Times New Roman", 10, True, False, False]}],
                }
                for style_id in ("NormalL1", "NormalL2", "NormalL3")
            ]
        }

        imported = docx_bytes_to_document(document_to_docx_bytes(document))
        styles = {style["id"]: style for style in imported["styles"]["paragraph"]}

        self.assertEqual(styles["NormalL1"]["outline_level"], 3)
        self.assertEqual(styles["NormalL2"]["outline_level"], 4)
        self.assertEqual(styles["NormalL3"]["outline_level"], 5)


if __name__ == "__main__":
    unittest.main()
