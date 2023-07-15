import pytest
from docx_charts.document import Document


@pytest.fixture
def doc():
    doc = Document('files/test.docx')
    yield doc
    doc.zipfs.close()


def test_tempdir(doc):
    assert doc.zipfs is not None

def test_list_contents(doc):
    assert {'[Content_Types].xml', 'docProps', '_rels', 'customXml', 'word'}.issubset(set(doc.list_contents()))
