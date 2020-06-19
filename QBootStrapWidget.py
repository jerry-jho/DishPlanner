from PySide2.QtWidgets import *
from PySide2.QtWebEngineWidgets import *

G_PREFIX = """
<!doctype html>
<html lang="zh-cn">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.5.0/dist/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
  </head>
  <body>
"""

G_PROFIX = """
    <script src="https://cdn.jsdelivr.net/npm/jquery@3.5.1/dist/jquery.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.0/dist/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  </body>
</html>
"""

class QBootStrapTableWidget(QWebEngineView):
    def __init__(self,row,col):
        QWebEngineView.__init__(self)
        self.h_headers = ['' for i in range(col)]
        self.v_headers = ['' for i in range(row)]
        self.d = {}
        self._row,self._col = row,col

    def setHorizontalHeaderLabels(self,l):
        self.h_headers = l
        self.buildHTML()

    def setVerticalHeaderLabels(self,l):
        self.v_headers = l
        self.buildHTML()

    def setText(self,r,c,t):
        self.d[(r,c)] = t.replace('\n','<br>')
        self.buildHTML()

    def buildHTML(self):
        html = """
<table class="table table-bordered">
  <thead class="thead-light">
    <tr>
      <th scope="col"></th>"""
        html += "\n".join(['<th scope="col">%s</th>' % x for x in self.h_headers])
        html += """
    </tr>
  </thead>
  <tbody>
        """
        for r in range(self._row):
            html += '<tr>\n<th scope="row">%s</th>' % self.v_headers[r]
            for c in range(self._col):
                html += '<td>%s</td>\n' % self.d.get((r,c),"")
            html += "</tr>\n"
        html += """
  </tbody>
</table>        
        """
        self.setHtml(G_PREFIX + html + G_PROFIX)
        
