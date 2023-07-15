import pytest
from docx_charts.chart import Chart
from fs.zipfs import ZipFS


@pytest.fixture
def chart():
	chart = Chart(ZipFS('files/PersonalizedReport_DraftV6.docx'), 'charts/chart1.xml', 'Chart 2')
	yield chart
	chart.file.close()


def test_chart_loading(chart):
	assert chart.file is not None
	assert chart.name == 'Chart 2'

def test_data(chart):
	assert chart.data() == [
		('WBMT1051 Wiskunde 2', 0.43),
		('CSE1205 Linear Algebra', 0.38),
		('CTB1002 Linear Algebra', 0.19)
	]
