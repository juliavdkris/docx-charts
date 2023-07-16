from fs.base import FS
from fs.zipfs import ZipFS
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
		return [
			list(zip(
				[cat.firstChild.nodeValue for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')
     				if cat.firstChild and type(cat.firstChild) == minidom.Text],
				[float(val.firstChild.nodeValue) for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v')
     				if val.firstChild and type(val.firstChild) == minidom.Text]
			))
			for series in dom.getElementsByTagName('c:ser')
		]



if __name__ == '__main__':
	chart = Chart(ZipFS('files/test/test.docx'), 'charts/chart1.xml', 'Chart 2')
	print(chart.data())
