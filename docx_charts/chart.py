from fs.base import FS
from fs.zipfs import ZipFS, WriteZipFS
from xml.dom import minidom
from typing import IO, Any


DataPoint = tuple[str, float]
Series = list[DataPoint]


class Chart:
	file: IO[Any]
	name: str

	def __init__(self, zipfs: FS, path: str, name: str):
		self.file = zipfs.open(f'word/{path}')
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
		chart.file.seek(0)
		return series

	def write_data(self, data: list[Series]) -> None:
		dom = minidom.parse(self.file)
		for series, series_data in zip(dom.getElementsByTagName('c:ser'), data):
			for cat, cat_data in zip(series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v'), series_data[0]):
				if cat.firstChild and type(cat.firstChild) == minidom.Text:
					cat.firstChild.replaceWholeText(cat_data)
			for val, val_data in zip(series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v'), series_data[1]):
				if val.firstChild and type(val.firstChild) == minidom.Text:
					val.firstChild.replaceWholeText(str(val_data))

		self.file.seek(0)
		self.file.truncate()
		self.file.write(dom.toxml())
		self.file.flush()



if __name__ == '__main__':
	chart = Chart(ZipFS('files/test/test2.docx'), 'charts/chart1.xml', 'Chart 2')
	data = chart.data()

	data[0][0] = (data[0][0][0], data[0][0][1] + 0.1)
	data[0][1] = (data[0][1][0], data[0][1][1] - 0.1)
	chart.write_data(data)

	print(chart.data())
