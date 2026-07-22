from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_highlight_replaces_font_toolbar_button():
    html = (ROOT / "static" / "index.html").read_text(encoding="utf-8")

    assert 'id="fontAdvancedToggle"' not in html
    assert 'id="highlightAdvancedToggle"' in html
    assert 'id="highlightAdvancedContent"' in html


def test_removed_font_panel_controls_are_null_safe():
    script = (ROOT / "static" / "app.js").read_text(encoding="utf-8")

    assert 'formatPainterBtn?.classList.add("is-active")' in script
    assert 'formatPainterBtn?.classList.remove("is-active")' in script
    assert 'clearFormatBtn?.addEventListener("click"' in script


def test_toolbar_height_tracks_expanded_panels_without_wheel_resize():
    script = (ROOT / "static" / "app.js").read_text(encoding="utf-8")

    assert 'function syncToolbarHeightToExpandedPanel()' in script
    assert 'toolbar-control-row")?.addEventListener("wheel"' not in script
    assert "syncToolbarHeightToExpandedPanel();" in script


def test_style_update_preserves_run_backgrounds():
    script = (ROOT / "static" / "app.js").read_text(encoding="utf-8")

    assert "const preservedBackground = normalizeBackgroundColor(" in script
    assert "descriptorWithoutBackground[5] = \"\";" in script
    assert "if (preservedBackground) run.style.backgroundColor = preservedBackground;" in script


def test_paragraph_style_prefers_live_editor_selection_over_cached_range():
    script = (ROOT / "static" / "app.js").read_text(encoding="utf-8")

    assert "if (!selectionInsideEditor()) {\n    restoreEditorSelection();\n  }" in script
    assert "editor.addEventListener(\"click\", () => {\n  // Capture synchronously" in script
    assert "let lastClickedParagraph = null;" in script
    assert "? [lastClickedParagraph]\n    : selectedBlockElements();" in script
