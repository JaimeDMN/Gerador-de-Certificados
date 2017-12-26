import os
from   reportlab.lib.enums     import TA_JUSTIFY,TA_CENTER,TA_RIGHT
from   reportlab.lib.pagesizes import letter
from   reportlab.platypus      import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from   reportlab.lib.styles    import getSampleStyleSheet, ParagraphStyle
from   reportlab.lib.units     import mm
from   PyPDF2                  import PdfFileWriter, PdfFileReader
from   openpyxl                import load_workbook

pastas = [x[0][2:] for x in os.walk(".")]
try:
    pastas.remove('Basico')
except:
    pass
try:
    pastas.remove(".ipynb_checkpoints")
except:
    pass
try:
    pastas.remove("tcl")
except:
    pass
try:
    pastas.remove('')
except:
    pass

for pasta in pastas:
    # TEXTO
    TEXTO      = open(os.path.join(pasta,"1-TEXTO.txt"),'r')
    TEXTO     = TEXTO.read()
    # INFO
    INFO      = open(os.path.join(pasta,"2-INFO.txt"),'r')
    Lista     = INFO.read().split("\n") 

    # Dentro do Arquivo
    wb = load_workbook(filename = os.path.join(pasta,"4-Lista.xlsx"))
    sheet_ranges = wb.active

    for i in range(2,len(sheet_ranges["A"])+1):
        Nome      = sheet_ranges["A{}".format(i)].value
        Email     = sheet_ranges["B{}".format(i)].value
        #if(Email==None):
        output = PdfFileWriter()
        Template = PdfFileReader(open(os.path.join("Basico", "Template.pdf"), 'rb'))

        doc = SimpleDocTemplate("TEMP1.pdf",pagesize=letter,
                                rightMargin=66,leftMargin=174,
                                topMargin=300,bottomMargin=18)
        Story=[]




        ptext = '<font name=Times-Roman size=15>' + TEXTO.format(Nome) +'</font>'

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY, leading=25))
        styles.add(ParagraphStyle(name= 'Center', alignment=TA_CENTER))
        styles.add(ParagraphStyle(name= 'Right', alignment=TA_RIGHT))
        styles.add(ParagraphStyle(name= 'Menos_C', alignment=TA_CENTER,spaceBefore=-3, spaceAfter=-3))
        styles.add(ParagraphStyle(name= 'Menos_J', alignment=TA_JUSTIFY,spaceBefore=0, spaceAfter=0))


        Story.append(Paragraph(ptext, styles["Justify"]))

        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))

        ptext = '<font name=Times-Roman size=15>' + Lista[0] +'</font>'
        Story.append(Paragraph(ptext, styles["Center"]))

        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))

        im = Image(os.path.join(pasta,"3-Assinatura.jpg"), 60*mm, 15*mm)
        Story.append(im)

        Story.append(Spacer(1, 12))

        ptext = '<font name=Times-Roman size=8>' + Lista[1] +'</font>'
        Story.append(Paragraph(ptext, styles["Menos_C"]))



        ptext = '<font name=Times-Roman size=8>' + Lista[2] +'</font>'
        Story.append(Paragraph(ptext, styles["Menos_C"]))



        ptext = '<font name=Times-Roman size=8>' + Lista[3] +'</font>'
        Story.append(Paragraph(ptext, styles["Menos_C"]))
        doc.build(Story)
        Escrita = PdfFileReader(open("TEMP1.pdf", 'rb'))
        page = Template.getPage(0)
        page.mergePage(Escrita.getPage(0))
        output.addPage(page)

        doc = SimpleDocTemplate("TEMP2.pdf",pagesize=letter,
                                rightMargin=60,leftMargin=160,
                                topMargin=35,bottomMargin=18)
        ptext = '<font name=Times-Roman size=11>' + Lista[4] + '</font>'
        Story.append(Paragraph(ptext, styles["Menos_J"]))
        ptext = '<font name=Times-Roman size=11>' + Lista[5] + '</font>'
        Story.append(Paragraph(ptext, styles["Menos_J"]))
        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))
        Story.append(Spacer(1, 12))

        ptext = '<font name=Times-Roman size=11>'+ Lista[6] +'</font>'
        Story.append(Paragraph(ptext, styles["Right"]))
        doc.build(Story)
        Escrita = PdfFileReader(open("TEMP2.pdf", 'rb'))
        page = Template.getPage(1)
        page.mergePage(Escrita.getPage(0))
        output.addPage(page)
        with open(os.path.join(pasta,Nome+".pdf"), 'wb') as f:
            output.write(f)
        