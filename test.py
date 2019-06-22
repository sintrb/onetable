# -*- coding: UTF-8 -*

if __name__ == '__main__':
    import os, time, subprocess
    from onetable import DefualtTable, HtmlRender, XlsRender, XlsxRender, PdfRender, CSVRender

    table = DefualtTable({'fontSize': 10, 'color': 'red'})
    table.setStyles(
        ('x', {'fontSize': 14}),
        ('red', {'color': 'red'}),
        ('title', {'color': 'black', 'font': u'宋体', 'fontSize': 16, 'align': 'center'})
    )
    table.writeNextRowCols([('Title1', 3, 'title')])
    table.pushStatus()
    #     table.writeCellWithSpan(row=0, col=3, rowspan=2, colspan=1, value='H', style='x')
    table.popStatus()
    table.writeNextRowCols([1, 2, 3])
    table.writeNextRowCols([3, 2, 1, 1], style='red')

    table.writeNextRowCols([(u'确认书', 8, 'title')])
    table.writeNextRowCols([u'这是一个长文本测试，有可能需要自动换行的哈哈', u'这是一个长文本测试，\n有可能需要自动换行的哈哈', 'x'], style='red')
    table.setWidths(*[100 for _ in range(8)])

    renders = [
        (HtmlRender, 'html'),
        (XlsRender, 'xls'),
        (XlsxRender, 'xlsx'),
        (PdfRender, 'pdf'),
        (CSVRender, 'csv'),
    ]
    testpath = '/tmp/'
    name = str(int(time.time()))
    for clz, sufix in renders:
        try:
            p = os.path.join(testpath, 'test%s.%s' % (name, sufix))
            print('test for %s with %s' % (sufix, p))
            r = clz()
            with open(p, 'wb') as out:
                r.save(tab=table, out=out)
            try:
                # mac
                subprocess.call(["open", p])
            except:
                # windows
                os.system(p)
        except:
            import traceback

            traceback.print_exc()
