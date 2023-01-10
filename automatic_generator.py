
import os
import webbrowser
from powerpoint import pegarTextoSlideShow

legenda = pegarTextoSlideShow()

html = "<html>\n\t<head>\n\t\t<title>Legenda</title>\n"
html += "\t\t<style>\n\t\t\t.text {\n\t\t\t\ttext-align: center;\n\t\t\t\tfont-family: 'Arial';\n\t\t\t\twidth: 1640px;\n\t\t\t\theight: 20vh;\n\t\t\t\tbackground: rgb(0, 0, 0, 0.8);\n\t\t\t\tcolor: white;\n\t\t\t}"
html += "\n\n\t\t\t.retangulo {\n\t\t\t\twidth: 100vw;\n\t\t\t\theight: 3vh;\n\t\t\t\tbackground: rgb(255, 255, 255);\n\t\t\t}\n\t\t</style>"
html += '\n\t\t<script src="static/js/jquery-3.6.0.min.js"></script>'
html += '\n\t\t<script src="static/js/textFit.min.js"></script>\n\t</head>'
html += '\n\t<body>'

for item in legenda:
    html += '\n\t\t<div class="text">'
    html += item + '</div>'
    html += '\n\t\t<div class="retangulo"></div>'



html += '\n\t\t<script>'
html += "\n\t\t\ttextFit($('div'), {alignVert: true, multiLine: true, minFontSize: 10, maxFontSize: 100});"
html += '\n\t\t</script>'

html += '\n\t</body>'
html += '\n</html>'

file = open('result.html', 'w', encoding='utf-8')
file.write(html)
file.close()

webbrowser.open('file://' + os.path.realpath('result.html'))

#print(html)