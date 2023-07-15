from docx_charts.document import Document


# @test('Document creates a temporary directory')
def test_tempdir():
    doc = Document('files/PersonalizedReport_DraftV6.docx')
    assert doc.file is not None
    assert doc.extracted is not None
