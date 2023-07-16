import io
from xml.dom import minidom


DataPoint = tuple[str, float]
Series = list[DataPoint]


class Chart:
	file: io.TextIOWrapper
	name: str

	def __init__(self, path: str, name: str):
		self.file = open(path, 'r+')
		self.name = name

	def __str__(self) -> str:
		return f'{self.name} ({self.file.name})'
	__repr__ = __str__


	def data(self) -> list[Series]:
		dom = minidom.parse(self.file)
		series = [
			list(zip(
				[cat.firstChild.nodeValue for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
					if cat.firstChild and type(cat.firstChild) == minidom.Text],
				[float(val.firstChild.nodeValue) for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
					if val.firstChild and type(val.firstChild) == minidom.Text]
			))
			for series in dom.getElementsByTagName('c:ser')
		]
		self.file.seek(0)
		return series

	def write_data(self, data: list[Series]) -> None:
		dom = minidom.parse(self.file)
		for series, series_data in zip(dom.getElementsByTagName('c:ser'), data):
			cats = [cat.firstChild for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
				if cat.firstChild and type(cat.firstChild) == minidom.Text]
			vals = [val.firstChild for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
				if val.firstChild and type(val.firstChild) == minidom.Text]

			for cat, val, (new_cat, new_val) in zip(cats, vals, series_data):
				cat.replaceWholeText(new_cat)
				val.replaceWholeText(str(new_val))

		self.file.seek(0)
		self.file.truncate()
		self.file.write(dom.toxml())
		self.file.flush()
		self.file.seek(0)



if __name__ == '__main__':
	chart = Chart('files/test/chart1.xml', 'Chart 2')
	data = chart.data()

	data[0][0] = (data[0][0][0], data[0][0][1] + 0.1)
	data[0][1] = (data[0][1][0], data[0][1][1] - 0.1)
	chart.write_data(data)

	print(chart.data())
