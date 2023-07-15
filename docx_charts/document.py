from fs.base import FS
from fs.zipfs import ZipFS
from xml.dom import minidom


class Document:
	zipfs: FS

	def __init__(self, file_path: str):
		self.zipfs = ZipFS(file_path)

	def list_contents(self):
		return self.zipfs.listdir('/')

	def list_charts(self) -> list[tuple[int, str]]:
		charts: list[tuple[int, str]] = []
		with self.zipfs.open('word/document.xml') as f:
			dom = minidom.parse(f)
			for node in dom.getElementsByTagName('wp:docPr'):
				if node.firstChild and type(node.firstChild) == minidom.Element and node.firstChild.getAttribute('xmlns:a') == 'http://schemas.openxmlformats.org/drawingml/2006/main':
					chart_id = int(node.getAttribute('id'))
					chart_name = node.getAttribute('name')
					charts.append((chart_id, chart_name))
		return charts

	def find_charts_by_name(self, name: str) -> list[tuple[int, str]]:
		return [chart for chart in self.list_charts() if chart[1] == name]



if __name__ == '__main__':
	doc = Document('files/test/test.docx')
	print(doc.list_contents())
	print(doc.list_charts())
