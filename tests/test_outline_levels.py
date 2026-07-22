import unittest
from io import BytesIO
from pathlib import Path

from docx import Document

from docx_io import _builtin_styles, docx_bytes_to_document, document_to_docx_bytes


class OutlineLevelRoundTripTests(unittest.TestCase):
    def test_style_update_preserves_extended_outline_level_for_paragraph_tags(self):
        script = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")

        self.assertIn("return style.outline_level ?? null;", script)

    def test_outline_items_use_fixed_level_indentation(self):
        css = (Path(__file__).resolve().parents[1] / "static" / "styles.css").read_text(encoding="utf-8")
        script = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")

        self.assertIn("padding: 5px 0;", css)
        self.assertNotIn(".outline-item.is-active {\n  padding-left:", css)
        self.assertIn("button.style.paddingLeft = `${indentLevel * 12}px`;", script)

    def test_all_outline_levels_support_folding(self):
        script = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")

        self.assertIn("function isOutlineHeading(element)", script)
        self.assertIn("const collapsedAncestors = [];", script)
        self.assertIn("while (collapsedAncestors.length && collapsedAncestors[collapsedAncestors.length - 1].level >= level)", script)
        self.assertIn("const headings = Array.from(editor.children).filter(isOutlineHeading);", script)

    def test_bulk_folding_batches_overlay_layout(self):
        script = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")

        self.assertIn("const placements = headings.map((heading) =>", script)
        self.assertIn("const fragment = document.createDocumentFragment();", script)
        self.assertIn("function scheduleChapterFoldOverlay(headings)", script)
        self.assertIn("applyChapterFolding({ deferOverlay: true, normalizeHeadings: false });", script)

    def test_other_toggle_limits_primary_outline_to_selected_siblings(self):
        script = (Path(__file__).resolve().parents[1] / "static" / "app.js").read_text(encoding="utf-8")

        self.assertIn("function primarySiblingBranchBounds(allItems, anchorElement)", script)
        self.assertIn("const parentLevel = anchorLevel - 1;", script)
        self.assertIn("return { start: parentStart, end: parentEnd, level: anchorLevel };", script)
        self.assertIn("&& displayOutlineLevel(item) === branchBounds.level", script)
        self.assertNotIn("const constrainedLevels = [];", script)

    def test_builtin_styles_are_ordered_by_outline_level_without_code(self):
        styles = _builtin_styles()

        self.assertEqual(
            [style["id"] for style in styles],
            ["Normal", "Heading1", "Heading2", "Heading3", "NormalL1", "NormalL2", "NormalL3"],
        )
        self.assertEqual(
            [style["name"] for style in styles],
            ["Normal", "L0", "L1", "L2", "L3", "L4", "L5"],
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

    def test_outline_styles_have_fixed_body_indent_in_docx(self):
        document = {"blocks": []}
        exported = Document(BytesIO(document_to_docx_bytes(document)))

        self.assertEqual(exported.styles["L0"].paragraph_format.left_indent.twips, 0)
        self.assertEqual(exported.styles["L1"].paragraph_format.left_indent.twips, 16 * 20)
        self.assertEqual(exported.styles["L2"].paragraph_format.left_indent.twips, 14 * 2 * 20)
        self.assertEqual(exported.styles["L3"].paragraph_format.left_indent.twips, 10 * 3 * 20)
        self.assertEqual(exported.styles["L4"].paragraph_format.left_indent.twips, 10 * 4 * 20)
        self.assertEqual(exported.styles["L5"].paragraph_format.left_indent.twips, 10 * 5 * 20)

    def test_editor_draws_outline_connectors_without_text_spaces(self):
        root = Path(__file__).resolve().parents[1]
        script = (root / "static" / "app.js").read_text(encoding="utf-8")
        css = (root / "static" / "styles.css").read_text(encoding="utf-8")

        self.assertIn('target.dataset.outlineIndent = String(outlineLevel);', script)
        self.assertIn("const fixedButtonLeft =", script)
        self.assertIn('.editor > [data-outline-indent]:not([data-outline-indent="0"])::before', css)
        self.assertIn("width: var(--outline-indent);", css)


if __name__ == "__main__":
    unittest.main()
