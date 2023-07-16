import os
import tempfile
import shutil
from xml.dom import minidom

from docx_charts.chart import Chart


class Document:
	file_path: str
	extracted: tempfile.TemporaryDirectory

	def __init__(self, file_path: str):
		self.file_path = file_path
		self.extracted = tempfile.TemporaryDirectory()
		shutil.unpack_archive(file_path, self.extracted.name, format='zip')

	def list_contents(self):
		return os.listdir(self.extracted.name)

	def list_charts(self) -> list[Chart]:
		charts: list[Chart] = []
		with open(os.path.join(self.extracted.name, 'word/document.xml')) as doc:
			with open(os.path.join(self.extracted.name, 'word/_rels/document.xml.rels')) as rels:
				doc_dom = minidom.parse(doc)
				rels_dom = minidom.parse(rels)
				for node in doc_dom.getElementsByTagName('c:chart'):
					relationship_id = node.getAttribute('r:id')
					relationship = [rel for rel in rels_dom.getElementsByTagName('Relationship') if rel.getAttribute('Id') == relationship_id][0]
					path = os.path.join(self.extracted.name, 'word', relationship.getAttribute('Target'))
					name = node.parentNode.parentNode.parentNode.getElementsByTagName('wp:docPr')[0].getAttribute('name')
					charts.append(Chart(path, name))
		return charts

	def find_charts_by_name(self, name: str) -> list[Chart]:
		return [chart for chart in self.list_charts() if chart.name == name]
