import io
from xml.dom import minidom


DataPoint = tuple[str, float]
Series = list[DataPoint]


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


	def data(self) -> list[Series]:
		'''
		Gets the internal table data the chart is based on.

		Returns:
			A list of Series objects representing the data for the chart.
		'''
		dom = minidom.parse(self.file)
		series = [
			list(zip(
				[cat.firstChild.nodeValue for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
					if cat.firstChild and isinstance(cat.firstChild, minidom.Text)],
				[float(val.firstChild.nodeValue) for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
					if val.firstChild and isinstance(val.firstChild, minidom.Text)]
			))
			for series in dom.getElementsByTagName('c:ser')
		]
		self.file.seek(0)
		return series


	def write_data(self, data: list[Series]) -> None:
		'''
		Overwrites the data of the chart.

		Args:
			data (list[Series]): The data to write to the chart.

		Note:
			The data must be in the same size and shape as the original data.
			(i.e. the same number of series, categories, and values)
		'''
		dom = minidom.parse(self.file)
		for series, series_data in zip(dom.getElementsByTagName('c:ser'), data):
			cats = [cat.firstChild for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
				if cat.firstChild and isinstance(cat.firstChild, minidom.Text)]
			vals = [val.firstChild for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
				if val.firstChild and isinstance(val.firstChild, minidom.Text)]

			for cat, val, (new_cat, new_val) in zip(cats, vals, series_data):
				cat.replaceWholeText(new_cat)
				val.replaceWholeText(str(new_val))

		self.file.seek(0)
		self.file.truncate()
		self.file.write(dom.toxml())
		self.file.flush()
		self.file.seek(0)
