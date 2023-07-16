import pytest
import os
from docx_charts import Document


@pytest.fixture(params=['word', 'libreoffice'])
def doc(request):
	return Document(f'files/PersonalizedReport_DraftV6_{request.param}.docx')


def test_tempdir(doc):
	assert doc.extracted is not None

def test_list_contents(doc):
	assert {os.path.basename(path) for path in doc.list_contents()} == {'[Content_Types].xml', 'docProps', '_rels', 'customXml', 'word'}

def test_list_charts(doc):
	charts = doc.list_charts()
	assert len(charts) == 6
	assert charts[0].name == 'Chart 2'
	assert 'word/charts/chart1.xml' in charts[0].file.name

def test_find_charts_by_name(doc):
	charts = doc.find_charts_by_name('Chart 2')
	assert len(charts) == 1
	assert charts[0].name == 'Chart 2'
	assert 'word/charts/chart1.xml' in charts[0].file.name
