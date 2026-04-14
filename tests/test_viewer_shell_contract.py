from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_summary_first_shell_and_accessibility_contract():
    index_html = (ROOT / "web-ui" / "index.html").read_text(encoding="utf-8")

    assert '<div class="workspace-shell">' in index_html
    assert '<aside class="side-panel side-panel--utility workspace-utility">' in index_html
    assert '<details id="detailViewerPanel" class="panel detail-viewer-panel workspace-detail">' in index_html
    assert index_html.index('<aside class="side-panel side-panel--utility workspace-utility">') < index_html.index(
        '<details id="detailViewerPanel" class="panel detail-viewer-panel workspace-detail">'
    )

    assert 'id="summaryPrevBtn" type="button" class="ghost small-btn" aria-label="Onceki ozet kaydi"' in index_html
    assert 'id="summaryNextBtn" type="button" class="ghost small-btn" aria-label="Sonraki ozet kaydi"' in index_html
    assert 'id="sourcePrevBtn" type="button" class="ghost small-btn" aria-label="Onceki source"' in index_html
    assert 'id="sourceNextBtn" type="button" class="ghost small-btn" aria-label="Sonraki source"' in index_html
    assert 'id="chartPrevBtn" type="button" class="ghost small-btn" aria-label="Onceki chart"' in index_html
    assert 'id="chartNextBtn" type="button" class="ghost small-btn" aria-label="Sonraki chart"' in index_html
    assert 'id="layerPrevBtn" type="button" class="ghost small-btn" aria-label="Onceki layer"' in index_html
    assert 'id="layerNextBtn" type="button" class="ghost small-btn" aria-label="Sonraki layer"' in index_html

    assert 'id="statusText" role="status" aria-live="polite" aria-atomic="true"' in index_html
    assert 'id="progressLabel" aria-live="polite" aria-atomic="true"' in index_html
    assert 'id="logBox" role="log" aria-live="polite" aria-atomic="false" aria-relevant="additions text"' in index_html
    assert 'label for="periodMaxInput" class="axis-control-label"' in index_html
    assert '<h1>Total Displacement Calculator</h1>' in index_html
    assert 'Kaynak sistemi' not in index_html
    assert '<h2>Viewer sources</h2>' in index_html
    assert 'Loaded sources' not in index_html
    assert 'id="summaryVariantSelect" aria-label="Gorunen varyant secimi"' in index_html
