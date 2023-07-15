from fs.base import FS
from fs.zipfs import ZipFS
from xml.dom import minidom

from docx_charts.chart import Chart


class Document:
	zipfs: FS

	def __init__(self, file_path: str):
		self.zipfs = ZipFS(file_path)

	def list_contents(self):
		return self.zipfs.listdir('/')

	def list_charts(self) -> list[Chart]:
		charts: list[Chart] = []
		with self.zipfs.open('word/document.xml') as doc:
			with self.zipfs.open('word/_rels/document.xml.rels') as rels:
				doc_dom = minidom.parse(doc)
				rels_dom = minidom.parse(rels)
				for node in doc_dom.getElementsByTagName('c:chart'):
					chart_rid = node.getAttribute('r:id')
					chart_path = [rel for rel in rels_dom.getElementsByTagName('Relationship') if rel.getAttribute('Id') == chart_rid][0].getAttribute('Target')
					chart_name = node.parentNode.parentNode.parentNode.getElementsByTagName('wp:docPr')[0].getAttribute('name')
					charts.append(Chart(self.zipfs, chart_path, chart_name))
		return charts

	def find_charts_by_name(self, name: str) -> list[Chart]:
		return [chart for chart in self.list_charts() if chart.name == name]



if __name__ == '__main__':
	doc = Document('files/test/test.docx')
	print(doc.list_contents())
	print(doc.list_charts())
