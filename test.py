# -*- coding: UTF-8 -*

if __name__ == '__main__':
    import os, time, subprocess
    from onetable import DefualtTable, gen_dataitemstable, HtmlRender, XlsRender, XlsxRender, PdfRender, CSVRender

    table = DefualtTable({'fontSize': 10, 'color': 'red'})
    cols = list(filter(lambda x: x, [
        {"field": "a", "title": u"团号"},
        {"field": "b", "title": u"酒店"},
    ]))
    items = [{'a': 'x', 'b': 'y'}, {'a': 'z', 'b': 't'}]
    table = gen_dataitemstable(items=items, table=DefualtTable(), columns=cols)

    renders = [
        (HtmlRender, 'html'),
        # (XlsRender, 'xls'),
        # (XlsxRender, 'xlsx'),
        # (PdfRender, 'pdf'),
        # (CSVRender, 'csv'),
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
