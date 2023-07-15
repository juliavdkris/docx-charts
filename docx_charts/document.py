import io
import tempfile
import shutil


class Document:
	file: io.IOBase
	extracted: tempfile.TemporaryDirectory

	def __init__(self, file_path: str):
		self.file = open(file_path, "rb")
		self.extracted = tempfile.TemporaryDirectory()
		shutil.unpack_archive(file_path, self.extracted.name, format='zip')
