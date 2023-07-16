import pytest
from docx_charts import Document, Chart


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
	assert chart.data() == [{
		'WBMT1051 Wiskunde 2': 0.43,
		'CSE1205 Linear Algebra': 0.38,
		'CTB1002 Linear Algebra': 0.19
	}]

def test_write_data(chart):
	chart.write_data([{
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.2,
		'CTB1002 Linear Algebra': 0.7
	}])
	assert chart.data() == [{
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.2,
		'CTB1002 Linear Algebra': 0.7
	}]

# Only overwrite the value of one category, the rest should remain unchanged
def test_write_data_only_one_cat(chart):
	chart.write_data([{
		'WBMT1051 Wiskunde 2': 0.1,
	}])
	assert chart.data() == [{
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.38,
		'CTB1002 Linear Algebra': 0.19
	}]

# The old implementation would fail when the new data dict has a different order than the original data
def test_write_data_out_of_order(chart):
	chart.write_data([{
		'CTB1002 Linear Algebra': 0.7,
		'CSE1205 Linear Algebra': 0.2,
		'WBMT1051 Wiskunde 2': 0.1,
	}])
	assert chart.data() == [{
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.2,
		'CTB1002 Linear Algebra': 0.7
	}]
