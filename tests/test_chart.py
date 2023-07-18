import pytest
from docx_charts import Document, Chart


@pytest.fixture(params=['word', 'libreoffice'])
def chart(request):
	doc = Document(f'files/PersonalizedReport_DraftV6_{request.param}.docx')
	chart = doc.charts_by_name('Chart 2')[0]
	yield chart
	chart.file.close()

@pytest.fixture(params=['word', 'libreoffice'])
def chart11(request):
	doc = Document(f'files/PersonalizedReport_DraftV6_{request.param}.docx')
	chart = doc.charts_by_name('Chart 11')[0]
	yield chart
	chart.file.close()


def test_chart_loading(chart):
	assert chart.file is not None
	assert chart.name == 'Chart 2'

def test_data(chart):
	assert chart.data() == {'Series1': {
		'WBMT1051 Wiskunde 2': 0.43,
		'CSE1205 Linear Algebra': 0.38,
		'CTB1002 Linear Algebra': 0.19
	}}

def test_data_multiple_series(chart11):
	assert chart11.data() == {
		'My score': {
			'Week2': 3,
			'Week7': 2.5
		},
		'Total score': {
			'Week2': 2.99,
			'Week7': 3.01
		}
	}

def test_write(chart):
	chart.write({'Series1': {
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.2,
		'CTB1002 Linear Algebra': 0.7
	}})
	assert chart.data() == {'Series1': {
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.2,
		'CTB1002 Linear Algebra': 0.7
	}}

# Only overwrite the value of one category, the rest should remain unchanged
def test_write_only_one_cat(chart):
	chart.write({'Series1': {
		'WBMT1051 Wiskunde 2': 0.1,
	}})
	assert chart.data() == {'Series1': {
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.38,
		'CTB1002 Linear Algebra': 0.19
	}}

# The old implementation would fail when the new data dict has a different order than the original data
def test_write_out_of_order(chart):
	chart.write({'Series1': {
		'CTB1002 Linear Algebra': 0.7,
		'CSE1205 Linear Algebra': 0.2,
		'WBMT1051 Wiskunde 2': 0.1,
	}})
	assert chart.data() == {'Series1': {
		'WBMT1051 Wiskunde 2': 0.1,
		'CSE1205 Linear Algebra': 0.2,
		'CTB1002 Linear Algebra': 0.7
	}}
