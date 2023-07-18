# docx-charts

Python library for manipulating chart data in word documents

Disclaimer: this is a project made for a very specific usecase, because nothing better existed yet. No stability guarantees are made, and the API is subject to change :)

&nbsp;

# Example
```py
from docx_charts import Document

doc = Document('example.docx')
chart = doc.find_charts_by_name('Chart 1')[0]
data = chart.data()
data['student_total']['CTB1002 Linear Algebra'] = 0.6
chart.write(data)
doc.save()
```
