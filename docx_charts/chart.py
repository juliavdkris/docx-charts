import io
from xml.dom import minidom


Series = dict[str, float|None]  # category: value
ChartData = dict[str, Series]  # series_name: Series


def series_name(series: minidom.Element) -> str | None:
	if not series.getElementsByTagName('c:tx'):
		return None
	return series.getElementsByTagName('c:tx')[0].getElementsByTagName('c:v')[0].firstChild.nodeValue  # type: ignore


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
		self.file = open(path, 'r+', encoding='utf-8')
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
			A dictionary of Series dicts representing data for the chart.
		'''
		dom = minidom.parse(self.file)
		series: ChartData = {
			series_name(series) or f'Series{i+1}':
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
			data (dict[str, dict[str, float|None]]): The data to write to the chart. The first key is the name of the series, the second key is the name of the category.

		Note:
			The names of the series and categories must be the same as in the original data.
			However, unchanged series and categories may be left out.
			If a value is None, it will be removed from the chart.
		'''
		dom = minidom.parse(self.file)

		for (new_series_name, new_series) in data.items():
			try:
				series = [series for i, series in enumerate(dom.getElementsByTagName('c:ser')) if (series_name(series) or f'Series{i+1}') == new_series_name][0]
			except IndexError:
				raise ValueError(f'Chart "{self.name}" does not contain a series named "{new_series_name}".')
			cats = [cat.firstChild for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
				if cat.firstChild and isinstance(cat.firstChild, minidom.Text)]
			vals = [val.firstChild for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
				if val.firstChild and isinstance(val.firstChild, minidom.Text)]

			# For every value in the new data, find the old value in the chart and overwrite it.
			for new_cat, new_val in new_series.items():
				if new_cat not in [cat.nodeValue for cat in cats]:
					raise ValueError(f'Chart "{self.name}" does not contain a category named "{new_cat}".')
				for cat, val in zip(cats, vals):
					if cat.nodeValue == new_cat:
						# Overwrite the value in the chart. Or if it is None, remove it from the chart.
						if new_val is None:
							val.parentNode.removeChild(val)
						else:
							val.replaceWholeText(str(new_val))
						break

		self.file.seek(0)
		self.file.truncate()
		self.file.write(dom.toxml())
		self.file.flush()
		self.file.seek(0)
