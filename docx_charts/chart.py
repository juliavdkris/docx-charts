import io
from xml.dom import minidom


Series = dict[str, float]  # category: value
ChartData = dict[str, Series]  # series_name: Series


def series_name(series: minidom.Element) -> str | None:
	if not series.getElementsByTagName('c:tx'):
		return None
	return series.getElementsByTagName('c:tx')[0].getElementsByTagName('c:v')[0].firstChild.nodeValue.lower().replace(' ', '_')  # type: ignore


class Chart:
	'''
	Represents a chart in a Word document.

	Attributes:
		file (io.TextIOWrapper): The file handler for the chart's XML.
		name (str): The name of the chart.
	'''
	file: io.TextIOWrapper
	name: str


	def __init__(self, path: str, name: str):
		'''
		Initializes a new instance of Chart.

		Args:
			path (str): The path to the chart's XML file.
			name (str): The name of the chart.
		'''
		self.file = open(path, 'r+')
		self.name = name


	def __del__(self):
		'''
		Closes the file handler when the Chart object is destroyed.
		'''
		self.file.close()


	def __str__(self) -> str:
		'''
		Returns a string representation of the Chart object.

		Returns:
			A string representation of the Chart object.
		'''
		return f'{self.name} ({self.file.name})'

	__repr__ = __str__


	def data(self) -> ChartData:
		'''
		Gets the internal table data the chart is based on.

		Returns:
			A list of Series dictionaries representing the data for the chart.
		'''
		dom = minidom.parse(self.file)
		series = {
			series_name(series) or f'series{i+1}':
			dict(zip(
				[cat.firstChild.nodeValue for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
					if cat.firstChild and isinstance(cat.firstChild, minidom.Text) and isinstance(cat.firstChild.nodeValue, str)],
				[float(val.firstChild.nodeValue) for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
					if val.firstChild and isinstance(val.firstChild, minidom.Text)]
			))
			for i, series in enumerate(dom.getElementsByTagName('c:ser'))
		}
		self.file.seek(0)
		return series


	def write(self, data: ChartData) -> None:
		'''
		Overwrites the data of the chart.

		Args:
			data (list[Series]): The data to write to the chart.

		Note:
			The data must be in the same number of series as the original data.
			However, unchanged categories may be left out.
		'''
		assert len(data) == len(self.data()), 'The new data must have the same number of series as the original data.'
		dom = minidom.parse(self.file)

		for (new_series_name, new_series) in data.items():
			series = [series for i, series in enumerate(dom.getElementsByTagName('c:ser')) if (series_name(series) or f'series{i+1}') == new_series_name][0]
			cats = [cat.firstChild for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
				if cat.firstChild and isinstance(cat.firstChild, minidom.Text)]
			vals = [val.firstChild for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
				if val.firstChild and isinstance(val.firstChild, minidom.Text)]

			for new_cat, new_val in new_series.items():
				assert new_cat in [cat.nodeValue for cat in cats], 'The new data must have the same categories as the original data.'
				for cat, val in zip(cats, vals):
					if cat.nodeValue == new_cat:
						val.replaceWholeText(str(new_val))
						break

		self.file.seek(0)
		self.file.truncate()
		self.file.write(dom.toxml())
		self.file.flush()
		self.file.seek(0)
