from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_chapter_fold_controls_stay_outside_editable_document():
    html = (ROOT / "static" / "index.html").read_text(encoding="utf-8")
    script = (ROOT / "static" / "app.js").read_text(encoding="utf-8")

    overlay_position = html.index('id="chapterFoldOverlay"')
    editor_position = html.index('id="editor"')

    assert overlay_position < editor_position
    assert 'heading.appendChild(button)' not in script
    assert 'chapterFoldOverlay.appendChild(button)' in script
    assert 'node.dataset.editorUi === "true"' in script


def test_deleted_heading_restores_the_following_paragraph_style():
    script = (ROOT / "static" / "app.js").read_text(encoding="utf-8")

    assert 'pendingDeleteStructure = (event.inputType || "").startsWith("delete")' in script
    assert "restoreStyleAfterDeletedHeading(deleteSnapshot)" in script
    assert 'currentText === nextText && currentText !== headingState.text' in script
