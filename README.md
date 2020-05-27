onetable
===============
Once defined, multiple ways to export table.

Install
===============
```
 pip install onetable
```

Useage
===============
```python
from onetable import DefualtTable, HtmlRender, XlsRender, XlsxRender, PdfRender, CSVRender

table = DefualtTable({'fontSize': 10, 'color': 'red'})
table.setStyles(
    ('x', {'fontSize': 14}),
    ('red', {'color': 'red'}),
    ('title', {'color': 'black', 'font': u'宋体', 'fontSize': 16, 'align': 'center'})
)
table.writeNextRowCols([('Title1', 3, 'title')])
table.pushStatus()
table.popStatus()
table.writeNextRowCols([1, 2, 3])
table.writeNextRowCols([3, 2, 1, 1], style='red')
with open('test.xls', 'wb') as out:
    XlsRender().save(tab=table, out=out)
```