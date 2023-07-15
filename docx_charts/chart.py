from fs.base import FS
from fs.zipfs import ZipFS
from xml.dom import minidom
from typing import IO, Any


class Chart:
	file: IO[Any]
	name: str

	def __init__(self, zipfs: FS, path: str, name: str):
		self.file = zipfs.open(f'word/{path}')
		self.name = name

	def __str__(self) -> str:
		return f'{self.name} ({self.file.name})'
	__repr__ = __str__


	def data(self) -> list[tuple[str, float]]:
		dom = minidom.parse(self.file)
		series = dom.getElementsByTagName('c:ser')[0]

		cats = [cat.firstChild.nodeValue for cat in series.getElementsByTagName('c:cat')[0].getElementsByTagName('c:v') if cat.firstChild and type(cat.firstChild) == minidom.Text]
		vals = [float(val.firstChild.nodeValue) for val in series.getElementsByTagName('c:val')[0].getElementsByTagName('c:v') if val.firstChild and type(val.firstChild) == minidom.Text]
		return list(zip(cats, vals))



if __name__ == '__main__':
	chart = Chart(ZipFS('files/test/test.docx'), 'charts/chart1.xml', 'Chart 2')
	print(chart.data())
