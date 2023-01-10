import win32com.client
import pythoncom

def pegarTextoSlideShow():   
    lista = []
    try:
        pp = win32com.client.GetActiveObject("PowerPoint.Application", pythoncom.CoInitialize())
    except:
        return None

    print(pp.Presentations.Count)
    if pp.Presentations.Count > 0:

    #filename = pp.SlideShowWindows(1).Presentation.Name
    #fullname = pp.SlideShowWindows(1).Presentation.FullName
    #index = pp.SlideShowWindows(1).View.Slide.SlideIndex

        for sld in pp.Presentations(pp.Presentations.Count).Slides:
            try:
                if sld.Shapes.Count > 0:

                    if sld.Shapes(1).HasTextFrame:
                        texto = sld.Shapes(1).TextFrame.TextRange.Text

                        if texto == '':
                            if sld.Shapes(2).HasTextFrame:
                                texto = sld.Shapes(2).TextFrame.TextRange.Text

                    elif sld.Shapes(2).HasTextFrame:
                        texto = sld.Shapes(2).TextFrame.TextRange.Text

                    else:
                        texto = '-'

                else:
                    texto = '-'
                
                texto = texto.replace(chr(11), ' ').replace(chr(13), '<br>').replace('  ', ' ')
                lista.append(texto)
            except:
                lista.append('-')

        return lista
        #return lista

    lista.append('-')
    return lista