import os
import tempfile
import shutil
from xml.dom import minidom

from docx_charts.chart import Chart


class Document:
	'''
	Represents a Word document.

	Attributes:
		file_path (str): The path to the Word document.
		extracted (tempfile.TemporaryDirectory): A temporary directory containing the extracted contents of the Word document.
	'''
	file_path: str
	extracted: tempfile.TemporaryDirectory


	def __init__(self, file_path: str):
		'''
		Initializes a new instance of Document.

		Args:
			file_path (str): The path to the Word document file.
		'''
		self.file_path = file_path
		self.extracted = tempfile.TemporaryDirectory()
		shutil.unpack_archive(file_path, self.extracted.name, format='zip')


	def __del__(self):
		'''
		Cleans up the temporary directory when the Document object is destroyed.
		'''
		self.extracted.cleanup()


	def save(self, file_path: str|None = None):
		'''
		Re-archives the Word document and saves it to the specified path.

		Args:
			file_path (str): The path to save the Word document to. If not specified, the original file will be overwritten.
		'''
		if file_path is None:
			file_path = self.file_path
		shutil.make_archive(file_path, 'zip', self.extracted.name)
		shutil.move(f'{file_path}.zip', file_path)


	def list_contents(self):
		'''
		Lists the contents of the Word document.

		Returns:
			A list of files in the root of the extracted archive
		'''
		return os.listdir(self.extracted.name)


	def charts(self) -> list[Chart]:
		'''
		Lists the charts in the Word document.

		Returns:
			A list of objects representing the charts in the Word document.
		'''
		charts: list[Chart] = []
		with open(os.path.join(self.extracted.name, 'word/document.xml'), encoding='utf-8') as doc:
			with open(os.path.join(self.extracted.name, 'word/_rels/document.xml.rels'), encoding='utf-8') as rels:
				doc_dom = minidom.parse(doc)
				rels_dom = minidom.parse(rels)
				for node in doc_dom.getElementsByTagName('c:chart'):
					relationship_id = node.getAttribute('r:id')
					relationship = [rel for rel in rels_dom.getElementsByTagName('Relationship') if rel.getAttribute('Id') == relationship_id][0]
					path = os.path.join(self.extracted.name, 'word', relationship.getAttribute('Target'))
					name = node.parentNode.parentNode.parentNode.getElementsByTagName('wp:docPr')[0].getAttribute('name')
					charts.append(Chart(path, name))
		return charts


	def charts_by_name(self, name: str) -> list[Chart]:
		'''
		Finds the charts in the Word document with the specified name.

		Args:
			name (str): The name of the charts to find.

		Returns:
			A list of objects representing the charts in the Word document with the specified name.

		Note:
			Generally this will only return one chart, but technically it's possible for there to be multiple charts with the same name.
		'''
		return [chart for chart in self.charts() if chart.name == name]
