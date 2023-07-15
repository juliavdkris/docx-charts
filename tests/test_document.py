import pytest
from docx_charts.document import Document


@pytest.fixture
def doc():
    doc = Document('files/test/test.docx')
    yield doc
    doc.zipfs.close()


def test_tempdir(doc):
    assert doc.zipfs is not None

def test_list_contents(doc):
    assert {'[Content_Types].xml', 'docProps', '_rels', 'customXml', 'word'}.issubset(set(doc.list_contents()))

def test_list_charts(doc):
    charts = doc.list_charts()
    assert len(charts) == 6
    assert charts[0].name == 'Chart 2'
    assert charts[0].file.name == 'word/charts/chart1.xml'

def test_find_charts_by_name(doc):
    charts = doc.find_charts_by_name('Chart 2')
    assert len(charts) == 1
    assert charts[0].name == 'Chart 2'
    assert charts[0].file.name == 'word/charts/chart1.xml'
