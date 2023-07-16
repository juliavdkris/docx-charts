import pytest
from docx_charts.document import Document
from docx_charts.chart import Chart


@pytest.fixture(params=['word', 'libreoffice'])
def chart(request):
	doc = Document(f'files/PersonalizedReport_DraftV6_{request.param}.docx')
	chart = doc.find_charts_by_name('Chart 2')[0]
	yield chart
	chart.file.close()


def test_chart_loading(chart):
	assert chart.file is not None
	assert chart.name == 'Chart 2'

def test_data(chart):
	assert chart.data() == [[
		('WBMT1051 Wiskunde 2', 0.43),
		('CSE1205 Linear Algebra', 0.38),
		('CTB1002 Linear Algebra', 0.19)
	]]
