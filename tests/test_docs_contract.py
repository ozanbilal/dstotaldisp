from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_docs_assets_and_links_exist():
    index_html = (ROOT / "web-ui" / "index.html").read_text(encoding="utf-8")
    docs_html_path = ROOT / "web-ui" / "docs.html"
    user_guide_path = ROOT / "docs" / "user-guide.md"
    developer_guide_path = ROOT / "docs" / "developer-guide.md"
    agents_path = ROOT / "AGENTS.md"
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    docs_html = docs_html_path.read_text(encoding="utf-8")
    user_guide = user_guide_path.read_text(encoding="utf-8")
    developer_guide = developer_guide_path.read_text(encoding="utf-8")
    agents_text = agents_path.read_text(encoding="utf-8")

    assert docs_html_path.exists()
    assert user_guide_path.exists()
    assert developer_guide_path.exists()
    assert agents_path.exists()

    assert 'href="./docs.html"' in index_html
    assert "../docs/user-guide.md" in docs_html
    assert "Arayuz akisi" in user_guide
    assert "varsayilan olarak kapali `Detayli inceleme` drawer" in user_guide
    assert "`Spectrum max period` yalniz o anda acik olan spectrum chart'ini sinirlar" in user_guide
    assert "`Viewer sources` sayaci dosya adedini degil" in user_guide
    assert "`shell`" in user_guide
    assert "docs/user-guide.md" in readme
    assert "docs/developer-guide.md" in readme
    assert "docs/developer-guide.md" in user_guide
    assert "Gelistirici / agent handoff dokumani `docs/developer-guide.md`" in agents_text
    assert "`sourceCatalog`" in developer_guide
    assert "`summaryCatalog`" in developer_guide
    assert "`/disp_core.py` ve `/web-ui/disp_core.py`" in developer_guide
    assert "`TIME_HISTORIES.LAYER#_DISP`" in developer_guide
    assert "geodisp/main" in developer_guide
    assert "Dokumantasyon guncellenmeden ozellik isi tamamlanmis sayilmaz." in agents_text
