from fs.base import FS
from fs.zipfs import ZipFS


class Document:
	zipfs: FS

	def __init__(self, file_path: str):
		self.zipfs = ZipFS(file_path)

	def list_contents(self):
		return self.zipfs.listdir('/')
