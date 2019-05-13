from django.http import HttpResponse, Http404, HttpResponseRedirect
from .models import Question, Choice
from django.template import loader
from django.shortcuts import render, get_object_or_404
from django.urls import reverse
from django.views import generic
from django.views.generic import TemplateView
from fpdf import FPDF
from win32com import client
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.utils import ImageReader

from tkinter import filedialog
import tkinter

#PDF libraries
from decimal import Decimal
from lxml import etree, objectify

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle
########
from flask import Flask, jsonify, request
from flask_cors import CORS
from flask_restful import Resource, Api

from collections import namedtuple

from unicodedata import normalize

import os, os.path
import requests
import xml.dom.minidom
import json
from xml.etree import ElementTree
import xlsxwriter
import reportlab
import datetime
import dateutil.parser
import code128
import string


app = Flask(__name__)
api = Api(app)
CORS(app)

# Variables Globales
fileUploaded = None
dataDocumentArray = []
tipoComprobanete = 'recibido'
file = None

def index(request):
    context = {'message': 'Hello world!', 'second': 'message 2!'}
    return render(request,'sri_test/index.html', context)

class sri_document(TemplateView):
    dataArray = ['something', 'here']
    
    
    
# app.run(host='10.0.2.15')
#@app.route('/read_txt', methods=['GET'])

def read_txt(request):
    control_flag = False
    xml_readed = open("recibidos.txt", "r",encoding="ISO-8859-1").readlines()
    # xml_string = f.read()
    # array = xml_string.split("\n")
    document_array = []
    obj = []
    for line in xml_readed:
        data = line.split("\t")
        if control_flag == False:
            for i in data:
                if i != "":
                    obj.append(i)
            control_flag = True
        else:
            obj.append(data[0])
            control_flag = False
            document_array.append(obj)
            obj = []
    # print(document_array)

    for i in document_array:
        # print("______________________________________")
        # print(i[8])
        headers = {'Content-Type': 'application/xml','Accept': 'application/xml'}
        body = "<Envelope xmlns=\"http://schemas.xmlsoap.org/soap/envelope/\">"
        body += "    <Body>"
        body += "       <autorizacionComprobante xmlns=\"http://ec.gob.sri.ws.autorizacion\">"
        body += "           <claveAccesoComprobante xmlns=\"\">"+i[8]+"</claveAccesoComprobante>"
        body += "       </autorizacionComprobante>"
        body += "    </Body>"
        body += "</Envelope>"
        r = requests.post(url="https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl", data=body, headers=headers)
        xml_response = r.text
        xml_response = xml_response.replace("&lt;","<")
        # xml = xml.dom.minidom.parseString(xml_response)
        # xml_pretty_str = xml.toprettyxml()
        f = open(i[8]+".xml","w+")
        f.write(xml_response)
    # print(aux2)
    # uglyxml = '<?xml version="1.0" encoding="UTF-8" ?><employees><employee><Name>Leonardo DiCaprio</Name></employee></employees>'
    # xml = xml.dom.minidom.parseString(uglyxml)
    # xml_pretty_str = xml.toprettyxml()
    # print(xml_pretty_str)
    context = {'second': document_array}
    return render(request,'sri_test/read_text.html', context)

#@app.route('/get_xml_sri/<string:auth_number>', methods=['GET'])
def get_xml_sri(auth_number):
    headers = {'Content-Type': 'application/xml','Accept': 'application/xml'}
    body = "<Envelope xmlns=\"http://schemas.xmlsoap.org/soap/envelope/\">"
    body += "    <Body>"
    body += "       <autorizacionComprobante xmlns=\"http://ec.gob.sri.ws.autorizacion\">"
    body += "           <claveAccesoComprobante xmlns=\"\">"+auth_number+"</claveAccesoComprobante>"
    body += "       </autorizacionComprobante>"
    body += "    </Body>"
    body += "</Envelope>"
    r = requests.post(url="https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl", data=body, headers=headers)
    xml_response = r.text
    xml_response = xml_response.replace("&lt;","<")
    xml_response = xml_response.replace("<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"><soap:Body><ns2:autorizacionComprobanteResponse xmlns:ns2=\"http://ec.gob.sri.ws.autorizacion\"><RespuestaAutorizacionComprobante><claveAccesoConsultada>"+auth_number+"</claveAccesoConsultada><numeroComprobantes>1</numeroComprobantes><autorizaciones>","<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>")
    xml_response = xml_response.replace("</autorizaciones></RespuestaAutorizacionComprobante></ns2:autorizacionComprobanteResponse></soap:Body></soap:Envelope>","")
    xml_response = xml_response.replace("<?xml version=\"1.0\" encoding=\"UTF-8\"?>","")
    f = open(auth_number+".xml","w+")
    f.write(xml_response)
    # make_pdf(xml_response,auth_number)
    return jsonify({"SRI":1})

def test(request):

    global fileUploaded, dataDocumentArray, file

    if request.method == 'POST' and len(request.FILES) != 0:

        control_flag = False
        uploaded_file = request.FILES['document']
        xml_readed2 = uploaded_file.readlines()
        file = xml_readed2

        #for line in xml_readed2:
            #data = b"line".decode("ISO-8859-1")
            #print(str(line, 'ISO-8859-1').split("\t"))
        
        document_array = []
        newArray = []
        obj = []
        for line in xml_readed2:
            data = str(line, 'ISO-8859-1').split("\t")
            if control_flag == False:
                for i in data:
                    if i != "":
                        obj.append(i)
                control_flag = True
            else:
                obj.append(data[0])
                control_flag = False
                document_array.append(obj)
                obj = []
            dataDocumentArray = document_array
        for item in document_array[1:]:
            newArray.append(item)

        context = {'comprobantes_data': newArray, 'arreglo': document_array}

        # print(aux2)
        # uglyxml = '<?xml version="1.0" encoding="UTF-8" ?><employees><employee><Name>Leonardo DiCaprio</Name></employee></employees>'
        # xml = xml.dom.minidom.parseString(uglyxml)
        # xml_pretty_str = xml.toprettyxml()
        # print(xml_pretty_str)
        #context = {'second': document_array}
        #return render(request,'sri_test/test.html', context)
        #return render(request, 'sri_test/test.html')
        fileUploaded = True    
        return render(request, 'sri_test/test.html', context)
        
    else:
        fileUploaded = False
        empty = []
        dataDocumentArray = empty
        return render(request, 'sri_test/test.html')

def downloadxml(request):
    global fileUploaded, dataDocumentArray

    if len(dataDocumentArray) != 0:
        root = tkinter.Tk()
        root.lift()
        root.attributes('-topmost',True)
        root.after_idle(root.attributes,'-topmost',False)
        root.geometry("0x0")
        dirname = filedialog.askdirectory(parent=root, initialdir="/", title='Please select a directory')
        print(dirname)
        root.destroy()
        if dirname != '':
            for i in dataDocumentArray[1:]:

                claveAcceso = ''
                if i[0] == 'Factura':
                    claveAcceso = i[8]
                if i[0] == 'Comprobante de Retención':
                    claveAcceso = i[9]
                if i[0] == 'Notas de Crédito':
                    claveAcceso = i[9]
                if i[0] == 'Notas de Débito':
                    claveAcceso = i[9]

                xml_response = bodyHeader2(claveAcceso)
                xml_response = xml_response.replace("&lt;","<")
                xml_response = xml_response.replace("&#xd;","")
                with open(os.path.join(os.path.join(os.path.expanduser('~'),dirname,i[1]+".xml")), "w+") as file1:
                    file1.write(xml_response)
                #f = open(i[8]+".xml","w+")
                #f.write(xml_response)
                #print(dataDocumentArray)
            return HttpResponse(1)
        else:
            return HttpResponse(2)
    else:
        return HttpResponse(0)
        
def downloadPdf(request):
    
    global fileUploaded, dataDocumentArray

    if len(dataDocumentArray) != 0 :
        root = tkinter.Tk()
        root.lift()
        root.attributes('-topmost',True)
        root.after_idle(root.attributes,'-topmost',False)
        root.geometry("0x0")
        dirname = filedialog.askdirectory(parent=root, initialdir="/", title='Please select a directory')
        print(dirname)
        root.destroy()

        #Eestilos para los párrafos que usaran en la tabla

        #Paragraph style
        style1 = ParagraphStyle('parrafo', fontName = "Helvetica-Bold", fontSize = 10 )
        style2 = ParagraphStyle('parrafo', fontName = "Helvetica-Bold", fontSize = 8 )
        style3 = ParagraphStyle('parrafo', fontName = "Helvetica-Bold", fontSize = 7, alignment = TA_CENTER, )
        style4 = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 8 )
        productsLeftStyle = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 7, alignment = TA_LEFT )
        productsCenterStyle = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 7, alignment = TA_CENTER, )
        productosRightStyle = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 7, alignment = TA_RIGHT, )


        if dirname != '':

            for i in dataDocumentArray[1:]:
        
                arrayData = bodyHeader(i)

                #Creación del PDF
                pathDir = dirname + '/'+ i[1] + '.pdf'
                print(pathDir)
                c = canvas.Canvas(pathDir, pagesize=A4)
                c.translate(0,(0.7)*inch)
                #c.rotate(180)

                #rect (w,h)
                c.roundRect(4*inch, (6.5)*inch, 287, 300, 4)
                c.roundRect((0.2)*inch, (6.5)*inch, 260, 200, 4)
                c.rect((0.2)*inch, (5.4)*inch, 560, 75)

                c.setFont("Helvetica", 8)
                c.setFillColorRGB(255,0,0)
                
                #Logo
                c.setFont("Helvetica-Bold", 28)
                c.setFillColorRGB(255,0,0)
                message = 'NO TIENE LOGO'
                c.drawString(0.5*inch, (10.1)*inch, message)

                #First Square
   
                p = Paragraph(arrayData[3], style1)
                p.wrapOn(c, (3.4)*inch, (2.5)*inch)  # size of 'textbox' for linebreaks etc.
                p.drawOn(c, (0.3)*inch, (8.9)*inch)

                p = Paragraph(arrayData[10], style2)
                p.wrapOn(c, (3.4)*inch, (2.5)*inch)
                p.drawOn(c, (0.3)*inch, (8.3)*inch)


                p = Paragraph('Dirección Matriz', productsLeftStyle)
                p1 = Paragraph('Dirección sucursal', productsLeftStyle) 
                p2 = Paragraph(arrayData[4], productsLeftStyle)   
                p3 = Paragraph(arrayData[11], productsLeftStyle)

                size = A4
                dataDirecciones = [[p, p2],[p1, p3]]
                tableDirecciones = Table(dataDirecciones, colWidths=[50, 200])
                tableDirecciones.canv = c
                w, heightAux = tableDirecciones.wrap(0,0)
                tableDirecciones.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ("ALIGN", (0,0), (0,-1), "LEFT"),
                                    ("ALIGN", (1,0), (1,-1), "RIGHT"),])
                tableDirecciones.wrapOn(c, size[0], size[1])
                tableDirecciones.drawOn(c, (0.3)*inch, (7.2)*inch)

                c.setFont("Helvetica", 8)
                c.setFillColorRGB(0,0,0)
                message = 'OBLIGADO A LLEVAR:                   ' + arrayData[13]
                c.drawString((0.3)*inch, (6.6)*inch, message)
                
                #Second Square

                c.setFont("Helvetica", 14)
                c.setFillColorRGB(0,0,0)
                message = 'R.U.C.: ' + arrayData[6]
                c.drawString((4.1)*inch, (10.3)*inch, message)
                c.drawString((4.1)*inch, (10)*inch, arrayData[8])

                c.setFont("Helvetica", 10)
                message = 'No.  ' + arrayData[9]
                c.drawString((4.1)*inch, (9.7)*inch, message)

                message = 'NÚMERO DE AUTORIZACIÓN'
                c.drawString((4.1)*inch, (9.4)*inch, message)
                c.setFont("Helvetica", 7)
                c.drawString((4.1)*inch, (9.1)*inch, arrayData[0])

                c.setFont("Helvetica", 10)
                message = 'FECHA Y HORA DE'
                c.drawString((4.1)*inch, (8.8)*inch, message)

                fechaAut = dateutil.parser.parse(arrayData[1])

                message = 'AUTORIZACIÓN:               ' + str(fechaAut)
                c.drawString((4.1)*inch, (8.6)*inch, message)

                message = 'AMBIENTE:                        ' + arrayData[2]
                c.drawString((4.1)*inch, (8.3)*inch, message)

                message = 'EMISIÓN:                           ' + arrayData[5]
                c.drawString((4.1)*inch, (8.0)*inch, message)

                message = 'CLAVE DE ACCESO'
                c.drawString((4.1)*inch, (7.7)*inch, message)

                logo = ImageReader(code128.image(arrayData[0]))
                c.drawImage(logo,(4.2)*inch, (6.9)*inch,  width=250, height=40)
                #message = 'HERE GOES THE IMAGE'
                #c.drawString((4.1)*inch, (6.9)*inch, message)

                c.setFont("Helvetica", 7)
                c.drawString((4.6)*inch, (6.7)*inch, arrayData[0])

                c.translate(0*inch, 0*inch)

                #Third square
                
                c.setFont("Helvetica", 8)
                c.setFillColorRGB(0,0,0)
                message = 'Razón Social/Nombres                                  '+ arrayData[14]
                c.drawString((0.3)*inch, (6.3)*inch, message)

                message = 'Identificacion                 '+ arrayData[15]
                c.drawString((0.3)*inch, (6.1)*inch, message)

                message = 'Fecha                            '+ arrayData[16] #+'                           Placa/Matrícula:                        '+'here goes lisence plate'
                c.drawString((0.3)*inch, (5.9)*inch, message)
            
                message = 'Dirección                       '+ arrayData[12]
                c.drawString((0.3)*inch, (5.7)*inch, message)

                # CREACION DE LA TABLA CON LOS PRODUCTOS

                #Campos de titulo
                p = Paragraph('Cod. Principal', style3)
                p1 = Paragraph('Cod. Auxiliar', style3)
                p2 = Paragraph('Cantidad', style3)
                p3 = Paragraph('Descripción', style3)
                p4 = Paragraph('Detalle Adicional', style3)
                p5 = Paragraph('Precio Unitario', style3)
                p6 = Paragraph('Subsidio', style3)
                p7 = Paragraph('Precio sin Subsidio', style3)
                p8 = Paragraph('Descuento', style3)
                p9 = Paragraph('Precio Total', style3)
                data = [[p, p1, p2, p3, p4, p5, p6,p7,p8, p9]]

                size = A4
                h = 0
                pagina = 0

                #Iteración del arreglo de datos para llenar la tabla
                for index, item in enumerate(arrayData[28]):
                    data.append(item)
                    table = Table(data, colWidths=[50, 50, 50, 80, 80, 50, 50, 50, 50, 50, 50])
                    table.canv = c
                    w, h = table.wrap(0,0)

                    #Primera hoja del pdf
                    if (h > 390 and pagina == 0 and index != 0 ):
                        pagina += 1
                        auxiliar = data
                        auxiliar.pop()

                        table = Table(auxiliar, colWidths=[50, 50, 50, 80, 80, 50, 50, 50, 50, 50, 50])
                        table.canv = c
                        table.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                                ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                ('BOX', (0,0), (-1,-1), 1, colors.black)])
                        table.canv = c
                        w, h = table.wrap(0,0)
                        table.wrapOn(c, size[0], size[1])
                        table.drawOn(c, (0.2)*inch, (390-h))

                        c.showPage()
                        c.translate(0,(0.7)*inch)
                        data = []
                        data.append(item)

                    # 2,3 ... hojas del pdf
                    if (h > 750):
                        auxiliar = data
                        auxiliar.pop()

                        table = Table(auxiliar, colWidths=[50, 50, 50, 80, 80, 50, 50, 50, 50, 50, 50])
                        table.canv = c
                        table.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                                ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                ('BOX', (0,0), (-1,-1), 1, colors.black)])
                        table.canv = c
                        w, h = table.wrap(0,0)
                        table.wrapOn(c, size[0], size[1])
                        table.drawOn(c, (0.2)*inch, (750-h))

                        c.showPage()
                        c.translate(0,(0.7)*inch)
                        data = []
                        data.append(item)

                #Crea la tabla de información adicional
                p = Paragraph('Información Adicional', style3)
                p1 = Paragraph(arrayData[29], style4)
                infoAdicionalArray = [[p], [p1]]
                tableAdicional = Table(infoAdicionalArray, colWidths=[300])
                tableAdicional.canv = c
                w, heightInfoAdicional = tableAdicional.wrap(0,0)
                tableAdicional.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ("ALIGN", (0,0), (-1,-1), "CENTER"),
                                    ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                    ('BOX', (0,0), (-1,-1), 1, colors.black)])
                tableAdicional.wrapOn(c, size[0], size[1])

                #Crea la tabla de forma de pago
                p = Paragraph('Forma de Pago', style3)
                p1 = Paragraph('Valor', style3)
                p2 = Paragraph(arrayData[18], productsLeftStyle)
                p3 = Paragraph(arrayData[19], productosRightStyle)
                formaPagoArray = [[p, p1], [p2, p3]]
                tablePago = Table(formaPagoArray, colWidths=[175, 75])
                tablePago.canv = c
                w, heightFormaPago = tablePago.wrap(0,0)
                tablePago.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                    ('BOX', (0,0), (-1,-1), 1, colors.black),
                                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                                    ('TEXTFONT', (0, 0), (-1, -1), 'Helvetica')])
                tablePago.wrapOn(c, size[0], size[1])

                #Crea la tabla de totales

                for item in arrayData[25]:
                    if item[0] == "iva":
                        iva = item[1]

                    if item[0] == "subDoce":
                        subtotal_doce = item[1]

                    if item[0] == "subCero":
                        subtotal_cero = item[1]

                    if item[0] == "noIva":
                        noObjetoIva = item[1]

                    if item[0] == "exeIva":
                        exentoIva = item[1]


                arregloTotales = [['SUBTOTAL 12%', subtotal_doce], 
                                ['SUBTOTAL 0%', subtotal_cero], 
                                ['SUBTOTAL NO OBJETO DE IVA', noObjetoIva], 
                                ['SUBTOTAL EXENTO DE IVA', exentoIva],
                                ['SUBTOTAL SIN IMPUESTOS', arrayData[20]],
                                ['TOTAL DESCUENTO', arrayData[17]],
                                ['ICE', arrayData[26]],
                                ['IVA 12%', iva],
                                ['IRBPNR', arrayData[27]],
                                ['PROPINA', arrayData[21]],
                                ['VALOR TOTAL', arrayData[22]]]

                tableTotal = Table(arregloTotales, colWidths=[150, 50])
                tableTotal.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ("ALIGN", (0,0), (0,-1), "LEFT"),
                                    ("ALIGN", (1,0), (1,-1), "RIGHT"),
                                    ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                    ('BOX', (0,0), (-1,-1), 1, colors.black),
                                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                                    ('TEXTFONT', (0, 0), (-1, -1), 'Helvetica')])
                tableTotal.wrapOn(c, size[0], size[1])

                #Crea la tabla de subsidios
                p = Paragraph('VALOR TOTAL SIN SUBSIDIO', productsLeftStyle)
                p1 = Paragraph('AHORRO POR SUBSIDIO', productsLeftStyle)
                p2 = Paragraph(str(arrayData[24]), productosRightStyle)
                p3 = Paragraph(arrayData[23], productosRightStyle)
                formaPagoArray = [[p, p2], [p1, p3]]
                tableSubsidio = Table(formaPagoArray, colWidths=[150, 50])
                tableSubsidio.canv = c
                w, heightSubsidio = tableSubsidio.wrap(0,0)
                tableSubsidio.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ("ALIGN", (0,0), (0,-1), "LEFT"),
                                    ("ALIGN", (1,0), (-1,-1), "RIGHT"),
                                    ('BOX', (0,0), (-1,-1), 1, colors.black)])
                tableSubsidio.wrapOn(c, size[0], size[1])

                tamanioTablaTotales = 0

                #Dibuja las tablas cuando alcanza en la primera hoja del pdf
                if (h <= 390 and pagina == 0):
                    table.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ("ALIGN", (0,0), (-1,-1), "CENTER"),
                                    ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                    ('BOX', (0,0), (-1,-1), 1, colors.black)])
                    table.wrapOn(c, size[0], size[1])
                    table.drawOn(c, (0.2)*inch, (390-h))

                    #Calculo del número de filas que entran en el espacio sobrante 
                    tamanioTablaTotales = (390-h) // 18

                    #La tabla de totales entra en el espacio de la primera página
                    if tamanioTablaTotales >= 11:
                        tableTotal.drawOn(c, (374.5), (390-h-198))

                        #La tabal de información adicional entra en el espacio sobrante
                        if heightInfoAdicional <= (385-h):
                            tableAdicional.drawOn(c, (0.2)*inch, (385-h-heightInfoAdicional))
                        
                            #Cuatro if's que determinan la posicion de las tablas de forma de pago y subsisio
                            if (heightFormaPago <= (380-h-heightInfoAdicional) and heightSubsidio <= (385-h-198)):
                                tablePago.drawOn(c, (0.2)*inch, (380-h-heightInfoAdicional-heightFormaPago))
                                tableSubsidio.drawOn(c, (374.5), (385-h-198-heightSubsidio))
                            
                            if (heightFormaPago > (380-h-heightInfoAdicional) and heightSubsidio > (385-h-198)):
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tablePago.drawOn(c, (0.2)*inch, (750-heightFormaPago))
                                tableSubsidio.drawOn(c, (374.5), (750-heightSubsidio))

                            if (heightFormaPago > (380-h-heightInfoAdicional) and heightSubsidio <= (385-h-198)):
                                tableSubsidio.drawOn(c, (374.5), (385-h-198-heightSubsidio))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tablePago.drawOn(c, (0.2)*inch, (750-heightFormaPago))
                            
                            if (heightFormaPago <= (380-h-heightInfoAdicional) and heightSubsidio > (385-h-198)):
                                tablePago.drawOn(c, (0.2)*inch, (380-h-heightInfoAdicional-heightFormaPago))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableSubsidio.drawOn(c, (374.5), (750-heightSubsidio))

                        #La tabal de información adicional no entra en el espacio sobrante
                        else:

                            #La tabla de subsidio entra en el espacio de la primera página y se
                            #dibujan las tablas de información adicional y forma de pago en otra página
                            if heightSubsidio <= (385-h-198):
                                tableSubsidio.drawOn(c, (374.5), (385-h-198-heightSubsidio))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                                tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))
                            
                            #Si dibujan las tres tablas (subsidios, totales e información adicional) 
                            # en una nueva página
                            else: 
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                                tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))
                                tableSubsidio.drawOn(c, (374.5), (750-heightSubsidio))

                    #Solo una parte de la tabla de totales entra en el espacio sobrante
                    if (tamanioTablaTotales < 11 and tamanioTablaTotales >= 1):

                        auxiliarTotales = []
                        auxiliarTotales2 = []

                        #Se crea dos nuevos arreglos con una porción del arreglo principal en cada uno,
                        #segun el espacio sobrante, para dibujarlos en páginas diferentes 
                        for item in arregloTotales[:(tamanioTablaTotales)]:
                            auxiliarTotales.append(item)
                        
                        for item in arregloTotales[tamanioTablaTotales:]:
                            auxiliarTotales2.append(item)
                        
                        #Se crea las tablas y estilos con los que serán dibujadas
                        tableTotal1 = Table(auxiliarTotales, colWidths=[150, 50])
                        tableTotal1.canv = c
                        w, heightAux = tableTotal1.wrap(0,0)
                        tableTotal1.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                            ("ALIGN", (0,0), (0,-1), "LEFT"),
                                            ("ALIGN", (1,0), (1,-1), "RIGHT"),
                                            ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                            ('BOX', (0,0), (-1,-1), 1, colors.black),
                                            ('FONTSIZE', (0, 0), (-1, -1), 7),
                                            ('TEXTFONT', (0, 0), (-1, -1), 'Helvetica')])
                        tableTotal1.wrapOn(c, size[0], size[1])
                        tableTotal1.drawOn(c, (374.5), (390-h-heightAux))

                        tableTotal2 = Table(auxiliarTotales2, colWidths=[150, 50])
                        tableTotal2.canv = c
                        w, heightAux2 = tableTotal2.wrap(0,0)
                        tableTotal2.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                            ("ALIGN", (0,0), (0,-1), "LEFT"),
                                            ("ALIGN", (1,0), (1,-1), "RIGHT"),
                                            ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                            ('BOX', (0,0), (-1,-1), 1, colors.black),
                                            ('FONTSIZE', (0, 0), (-1, -1), 7),
                                            ('TEXTFONT', (0, 0), (-1, -1), 'Helvetica')])
                        tableTotal2.wrapOn(c, size[0], size[1])

                        #La tabla de información adicional entra en el espacio sobrante de la primera página
                        if heightInfoAdicional <= (385-h):
                            tableAdicional.drawOn(c, (0.2)*inch, (385-h-heightInfoAdicional))

                            #La tabla de forma de pago entra en el espacio sobrante de la primera página y se dibuja 
                            #en una nueva página las tablas de totales y subsidio
                            if (heightFormaPago <= (380-h-heightInfoAdicional)):
                                tablePago.drawOn(c, (0.2)*inch, (380-h-heightInfoAdicional-heightFormaPago))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableTotal2.drawOn(c, (374.5), (750-heightAux2))
                                tableSubsidio.drawOn(c, (374.5), (745-heightAux2-heightSubsidio))
                            
                            #Las tres tablas se dibujan en una nueva página 
                            else:
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tablePago.drawOn(c, (0.2)*inch, (750-heightFormaPago))
                                tableTotal2.drawOn(c, (374.5), (750-heightAux2))
                                tableSubsidio.drawOn(c, (374.5), (745-heightAux2-heightSubsidio))

                        #Se dibuja las restantes tablas en una nueva página ya que ninguna entra en el epsacio sobrante
                        #de la primera página
                        else:
                            c.showPage()
                            c.translate(0,(0.7)*inch)
                            tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                            tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))
                            tableTotal2.drawOn(c, (374.5), (750-heightAux2))
                            tableSubsidio.drawOn(c, (374.5), (745-heightAux2-heightSubsidio))

                    #Ninguna tabla entra en el espacio sobrante
                    if tamanioTablaTotales <= 0:
                        c.showPage()
                        c.translate(0,(0.7)*inch)
                        tableTotal.drawOn(c, (374.5), (750-198))
                        tableSubsidio.drawOn(c, (374.5), (745-198-heightSubsidio))
                        tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                        tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))

                #Dibuja las tablas en las demas hojas del pdf
                if(h <= 750 and pagina != 0):
                    table.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                    ("ALIGN", (0,0), (-1,-1), "CENTER"),
                                    ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                    ('BOX', (0,0), (-1,-1), 1, colors.black)])
                    table.wrapOn(c, size[0], size[1])
                    table.drawOn(c, (0.2)*inch, (750-h))
                
                    #Calculo del número de filas que entran en el espacio sobrante 
                    tamanioTablaTotales = (750-h) // 18

                    #La tabla de totales entra en el espacio de la página
                    if tamanioTablaTotales >= 11:
                        tableTotal.drawOn(c, (374.5), (750-h-198))

                        #La tabla de información adicional entra en el espacio sobrante
                        if heightInfoAdicional <= (745-h):
                            tableAdicional.drawOn(c, (0.2)*inch, (745-h-heightInfoAdicional))
                        
                            #Cuatro if's que determinan la posicion de las tablas de forma de pago y subsisio
                            if (heightFormaPago <= (740-h-heightInfoAdicional) and heightSubsidio <= (745-h-198)):
                                tablePago.drawOn(c, (0.2)*inch, (740-h-heightInfoAdicional-heightFormaPago))
                                tableSubsidio.drawOn(c, (374.5), (745-h-198-heightSubsidio))
                            
                            if (heightFormaPago > (740-h-heightInfoAdicional) and heightSubsidio > (745-h-198)):
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tablePago.drawOn(c, (0.2)*inch, (750-heightFormaPago))
                                tableSubsidio.drawOn(c, (374.5), (750-heightSubsidio))

                            if (heightFormaPago > (740-h-heightInfoAdicional) and heightSubsidio <= (745-h-198)):
                                tableSubsidio.drawOn(c, (374.5), (745-h-198-heightSubsidio))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tablePago.drawOn(c, (0.2)*inch, (750-heightFormaPago))
                            
                            if (heightFormaPago <= (740-h-heightInfoAdicional) and heightSubsidio > (745-h-198)):
                                tablePago.drawOn(c, (0.2)*inch, (740-h-heightInfoAdicional-heightFormaPago))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableSubsidio.drawOn(c, (374.5), (750-heightSubsidio))

                        #La tabal de información adicional no entra en el espacio sobrante
                        else:

                            #La tabla de subsidio entra en el espacio de la página y se
                            #dibujan las tablas de información adicional y forma de pago en otra página
                            if heightSubsidio <= (745-h-198):
                                tableSubsidio.drawOn(c, (374.5), (745-h-198-heightSubsidio))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                                tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))
                            
                            #Si dibujan las tres tablas (subsidios, totales e información adicional) 
                            # en una nueva página
                            else: 
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                                tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))
                                tableSubsidio.drawOn(c, (374.5), (750-heightSubsidio))

                    #Solo una parte de la tabla de totales entra en el espacio sobrante
                    if (tamanioTablaTotales < 11 and tamanioTablaTotales >= 1):

                        auxiliarTotales = []
                        auxiliarTotales2 = []

                        #Se crea dos nuevos arreglos con una porción del arreglo principal en cada uno,
                        #segun el espacio sobrante, para dibujarlos en páginas diferentes 
                        for item in arregloTotales[:(tamanioTablaTotales)]:
                            auxiliarTotales.append(item)
                        
                        for item in arregloTotales[tamanioTablaTotales:]:
                            auxiliarTotales2.append(item)
                        
                        #Se crea las tablas y estilos con los que serán dibujadas
                        tableTotal1 = Table(auxiliarTotales, colWidths=[150, 50])
                        tableTotal1.canv = c
                        w, heightAux = tableTotal1.wrap(0,0)
                        tableTotal1.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                            ("ALIGN", (0,0), (0,-1), "LEFT"),
                                            ("ALIGN", (1,0), (1,-1), "RIGHT"),
                                            ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                            ('BOX', (0,0), (-1,-1), 1, colors.black),
                                            ('FONTSIZE', (0, 0), (-1, -1), 7),
                                            ('TEXTFONT', (0, 0), (-1, -1), 'Helvetica')])
                        tableTotal1.wrapOn(c, size[0], size[1])
                        tableTotal1.drawOn(c, (374.5), (750-h-heightAux))

                        tableTotal2 = Table(auxiliarTotales2, colWidths=[150, 50])
                        tableTotal2.canv = c
                        w, heightAux2 = tableTotal2.wrap(0,0)
                        tableTotal2.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                            ("ALIGN", (0,0), (0,-1), "LEFT"),
                                            ("ALIGN", (1,0), (1,-1), "RIGHT"),
                                            ('INNERGRID', (0,0), (-1,-1), 1, colors.black),
                                            ('BOX', (0,0), (-1,-1), 1, colors.black),
                                            ('FONTSIZE', (0, 0), (-1, -1), 7),
                                            ('TEXTFONT', (0, 0), (-1, -1), 'Helvetica')])
                        tableTotal2.wrapOn(c, size[0], size[1])

                        #La tabla de información adicional entra en el espacio sobrante de la primera página
                        if heightInfoAdicional <= (745-h):
                            tableAdicional.drawOn(c, (0.2)*inch, (745-h-heightInfoAdicional))

                            #La tabla de forma de pago entra en el espacio sobrante de la primera página y se dibuja 
                            #en una nueva página las tablas de totales y subsidio
                            if (heightFormaPago <= (740-h-heightInfoAdicional)):
                                tablePago.drawOn(c, (0.2)*inch, (740-h-heightInfoAdicional-heightFormaPago))
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tableTotal2.drawOn(c, (374.5), (750-heightAux2))
                                tableSubsidio.drawOn(c, (374.5), (745-heightAux2-heightSubsidio))
                            
                            #Las tres tablas se dibujan en una nueva página 
                            else:
                                c.showPage()
                                c.translate(0,(0.7)*inch)
                                tablePago.drawOn(c, (0.2)*inch, (750-heightFormaPago))
                                tableTotal2.drawOn(c, (374.5), (750-heightAux2))
                                tableSubsidio.drawOn(c, (374.5), (745-heightAux2-heightSubsidio))

                        #Se dibuja las restantes tablas en una nueva página ya que ninguna entra en el epsacio sobrante
                        #de la primera
                        else:
                            c.showPage()
                            c.translate(0,(0.7)*inch)
                            tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                            tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))
                            tableTotal2.drawOn(c, (374.5), (750-heightAux2))
                            tableSubsidio.drawOn(c, (374.5), (745-heightAux2-heightSubsidio))

                    #Ninguna tabla entra en el espacio sobrante
                    if tamanioTablaTotales <= 0:
                        c.showPage()
                        c.translate(0,(0.7)*inch)
                        tableTotal.drawOn(c, (374.5), (750-198))
                        tableSubsidio.drawOn(c, (374.5), (745-198-heightSubsidio))
                        tableAdicional.drawOn(c, (0.2)*inch, (750-heightInfoAdicional))
                        tablePago.drawOn(c, (0.2)*inch, (745-heightInfoAdicional-heightFormaPago))


                c.save()

            return HttpResponse(1)
        else:
            return HttpResponse(2)
    
    else:
        return HttpResponse(0)

def comprobantesRecibidos(request):
    global tipoComprobanete
    tipoComprobanete = 'recibido'
    print(tipoComprobanete)
    return HttpResponse('DONE')

def comprobantesEmitidos(request):
    global tipoComprobanete
    tipoComprobanete = 'emitido'
    print(tipoComprobanete)
    return HttpResponse('DONE')

def downloadeExcel(request):

    global fileUploaded, dataDocumentArray

    if len(dataDocumentArray) != 0 :
        
        root = tkinter.Tk()
        root.lift()
        root.attributes('-topmost',True)
        root.after_idle(root.attributes,'-topmost',False)
        dirname = filedialog.asksaveasfilename(filetypes = (("Excel files", "*.xlsx"),("All files", "*.*") ))
        workbook = xlsxwriter.Workbook(dirname+'.xlsx') 
        worksheet = workbook.add_worksheet() 
        print(dirname)
        root.destroy()
        if dirname != '':

            #Formats
            data_format = workbook.add_format({'text_wrap': True, 'font_size': 8, 'border': 1, 'valign': 'vcenter'})
            titles_format = workbook.add_format({'text_wrap': True, 'bold': 1, 'align': 'center', 'valign': 'vcenter','font_size': 10, 'border': 1,})
            merge_format = workbook.add_format({'text_wrap': True, 'bold': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12, 'border': 1,})
            
            #Tamaños de las columnas
            worksheet.set_column('A:A',20)
            worksheet.set_column('B:C',20)
            worksheet.set_column('D:D',20)
            worksheet.set_column('E:E',15)
            worksheet.set_column('G:G',30)
            worksheet.set_column('H:J',30)
            worksheet.set_column('K:K',20)
            worksheet.set_column('L:L',30)
            worksheet.set_column('M:M',20)
            worksheet.set_column('N:N',10)
            worksheet.set_column('O:O',30)
            worksheet.set_column('P:AB',15)
            worksheet.set_column('AC:AC',30)
            worksheet.set_column('AD:AE',30)

            #Tipo de campos
            worksheet.merge_range('A1:K1', 'Información Empresa', merge_format)
            worksheet.merge_range('L1:AB1', 'Información Factura', merge_format)
            worksheet.write('AC1', 'Información Adicional', merge_format)
            worksheet.merge_range('AD1:AE1', 'Forma de Pago', merge_format)

            #Capos descripcion
            worksheet.write('A2', 'Tipo de Comprobante', titles_format) 
            worksheet.write('B2', 'R.U.C', titles_format) 
            worksheet.write('C2', 'No', titles_format) 
            worksheet.write('D2', 'Fecha y Hora de Autorización', titles_format)
            worksheet.write('E2', 'Ambiente', titles_format)
            worksheet.write('F2', 'Emisión', titles_format)
            worksheet.write('G2', 'Clave de Acceso', titles_format)
            worksheet.write('H2', 'Nombre', titles_format)
            worksheet.write('I2', 'Dirección Matriz', titles_format)
            worksheet.write('J2', 'Dirección Sucursal', titles_format)
            worksheet.write('K2', 'Obligado a Llevar', titles_format)
            
            worksheet.write('L2', 'Razón Social/Nombre', titles_format)
            worksheet.write('M2', 'Identificación', titles_format)
            worksheet.write('N2', 'Fecha', titles_format)
            worksheet.write('O2', 'Dirección', titles_format)
            worksheet.write('P2', 'Subtotal 12%', titles_format)
            worksheet.write('Q2', 'Subtotal 0%', titles_format)
            worksheet.write('R2', 'Subtotal No Objeto de I.V.A', titles_format)
            worksheet.write('S2', 'Subtotal Exento I.V.A', titles_format)
            worksheet.write('T2', 'Subtotal Sin Impuestos', titles_format)
            worksheet.write('U2', 'Total Descuento', titles_format)
            worksheet.write('V2', 'ICE', titles_format)
            worksheet.write('W2', 'IVA 12%', titles_format)
            worksheet.write('X2', 'IRBPNR', titles_format)
            worksheet.write('Y2', 'Propina', titles_format)
            worksheet.write('Z2', 'Valor Total', titles_format)
            worksheet.write('AA2', 'Valor Total Sin Subsidio', titles_format)
            worksheet.write('AB2', 'Ahorro por Subsidio', titles_format)

            worksheet.write('AC2', 'Información Adicional', titles_format)

            worksheet.write('AD2', 'Forma de Pago', titles_format)
            worksheet.write('AE2', 'Valor', titles_format)

            #Extraer datos

            count = 2

            for i in dataDocumentArray[1:]:
                count += 1
                
                claveAcceso = ''
                if i[0] == 'Factura':
                    claveAcceso = i[8]
                if i[0] == 'Comprobante de Retención':
                    claveAcceso = i[9]
                if i[0] == 'Notas de Crédito':
                    claveAcceso = i[9]
                if i[0] == 'Notas de Débito':
                    claveAcceso = i[9]

                xml_response = bodyHeader2(claveAcceso)
                
                ns = {'soap':'http://schemas.xmlsoap.org/soap/envelope/'}
                ns2 = {'ns2': 'http://ec.gob.sri.ws.autorizacion'}
                root = ElementTree.fromstring(xml_response)

                numeroAutorizacion = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('numeroAutorizacion').text
                worksheet.write('G'+str(count), numeroAutorizacion, data_format)

                fechaAutorizacion = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('fechaAutorizacion').text
                worksheet.write('D'+str(count), fechaAutorizacion, data_format)
                
                ambiente = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('ambiente').text
                worksheet.write('E'+str(count), ambiente, data_format)
                
                claveAccesoConsultada = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('claveAccesoConsultada').text
                
                value = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('comprobante')
                
                if value is not None:

                    text2 = value.text.replace("&lt;","<")
                    rootFactura = ElementTree.fromstring(text2)
                    detalles = rootFactura.find('detalles')
                    infoTributaria = rootFactura.find('infoTributaria')
                    infoFactura = rootFactura.find('infoFactura')
                    infoAdicional = rootFactura.find('infoAdicional')

                    if infoTributaria is not None:

                        if infoTributaria.find('tipoEmision') is not None:
                            emisionNumero = int(infoTributaria.find('tipoEmision').text)
                            if emisionNumero == 1:
                                tipoEmision = 'NORMAL'
                            else:
                                tipoEmision = ''
                        else:
                            tipoEmision = ''
                        worksheet.write('F'+str(count), tipoEmision, data_format)
                        

                        if infoTributaria.find('ruc') is not None:
                            ruc = infoTributaria.find('ruc').text
                        else:
                            ruc = ''
                        worksheet.write('B'+str(count), ruc, data_format) 

                        if infoTributaria.find('razonSocial') is not None:
                            razonSocial = infoTributaria.find('razonSocial').text
                        else: 
                            razonSocial = ''
                        worksheet.write('H'+str(count), razonSocial, data_format)
                        
                        if infoTributaria.find('dirMatriz') is not None:
                            dirMatriz = infoTributaria.find('dirMatriz').text
                        else:
                            dirMatriz = ''
                        worksheet.write('I'+str(count), dirMatriz, data_format)

                        if (infoTributaria.find('estab') is not None and
                            infoTributaria.find('ptoEmi') is not None and
                            infoTributaria.find('secuencial') is not None):
                            estab = infoTributaria.find('estab').text
                            ptoEmi = infoTributaria.find('ptoEmi').text
                            secuencial = infoTributaria.find('secuencial').text
                            No = estab + '-' + ptoEmi + '-' + secuencial 
                        else:
                            No = ''  
                        worksheet.write('C'+str(count), No , data_format)    

                        if infoTributaria.find('nombreComercial') is not None:
                            nombreComercial = infoTributaria.find('nombreComercial').text
                        else:
                            nombreComercial = ''

                        if infoTributaria.find('codDoc') is not None:
                            codDoc = int(infoTributaria.find('codDoc').text)
                        else:
                            codDoc = 0
                            
                        if codDoc == 1:
                            codigDoc = 'FACTURA'
                        if codDoc == 4:
                            codigDoc = 'NOTA DE CRÉDITO'
                        if codDoc == 5:
                            codigDoc = 'NOTA DE DÉBITO'
                        if codDoc == 6:
                            codigDoc = 'GUÍA DE REMISIÓN'
                        if codDoc == 7:
                            codigDoc = 'COMPROBANTE DE RETENCIÓN'
                        if codDoc == 0:
                            codigDoc = ''
                        worksheet.write('A'+str(count), codigDoc, data_format)   
                    
                    if infoFactura is not None:

                        if infoFactura.find('dirEstablecimiento') is not None:
                            establecimiento = infoFactura.find('dirEstablecimiento').text
                            if establecimiento != dirMatriz:
                                dirEstablecimiento = infoFactura.find('dirEstablecimiento').text
                            else:
                                dirEstablecimiento = ''
                        else:
                            dirEstablecimiento = ''
                        worksheet.write('J'+str(count), dirEstablecimiento, data_format)

                        if infoFactura.find('direccionComprador') is not None: 
                            direccionComprador = infoFactura.find('direccionComprador').text
                        else:
                            direccionComprador = ''
                        worksheet.write('O'+str(count), direccionComprador, data_format)

                        if infoFactura.find('obligadoContabilidad') is not None: 
                            obligadoContabilidad = infoFactura.find('obligadoContabilidad').text
                        else:
                            obligadoContabilidad = ''
                        worksheet.write('K'+str(count), obligadoContabilidad, data_format)

                        if infoFactura.find('razonSocialComprador') is not None:
                            razonSocialComprador = infoFactura.find('razonSocialComprador').text
                        else:
                            razonSocialComprador = ''
                        worksheet.write('L'+str(count), razonSocialComprador, data_format)

                        if infoFactura.find('identificacionComprador') is not None: 
                            identificacionComprador = infoFactura.find('identificacionComprador').text
                        else:
                            identificacionComprador = ''
                        worksheet.write('M'+str(count), identificacionComprador, data_format)

                        if infoFactura.find('fechaEmision') is not None:
                            fechaEmision = infoFactura.find('fechaEmision').text
                        else:
                            fechaEmision = ''
                        worksheet.write('N'+str(count), fechaEmision, data_format)

                        if infoFactura.find('totalDescuento') is not None:
                            totalDescuento = infoFactura.find('totalDescuento').text
                        else:
                                totalDescuento = '0'
                        worksheet.write('U'+str(count), totalDescuento, data_format)

                        if infoFactura.find('pagos').find('pago').find('formaPago') is not None:
                            formaPagoNumero = int(infoFactura.find('pagos').find('pago').find('formaPago').text)
                            if formaPagoNumero == 1:
                                formaPago = 'SIN UTILIZACION DEL SISTEMA FINANCIERO'
                            elif formaPagoNumero == 15:
                                formaPago = 'COMPENSACIÓN DE DEUDAS'
                            elif formaPagoNumero == 16:
                                formaPago = 'TARJETA DE DÉBITO'
                            elif formaPagoNumero == 17:
                                formaPago = 'DINERO ELECTRÓNICO'
                            elif formaPagoNumero == 18:
                                formaPago = 'TARJETA PREPAGO'
                            elif formaPagoNumero == 19:
                                formaPago = 'TARJETA DE CRÉDITO'
                            elif formaPagoNumero == 20:
                                formaPago = 'OTROS CON UTILIZACION DEL SISTEMA FINANCIERO'
                            elif formaPagoNumero == 21:
                                formaPago = 'ENDOSO DE TÍTULOS'
                            else:
                                formaPago = ''
                        else:
                            formaPago = ''
                        worksheet.write('AD'+str(count), formaPago, data_format)

                        if infoFactura.find('pagos').find('pago').find('total') is not None:
                            total = infoFactura.find('pagos').find('pago').find('total').text
                        else:
                            total = '0' 
                        worksheet.write('AE'+str(count), total, data_format)

                        if infoFactura.find('totalSinImpuestos') is not None:
                                totalSinImpuestos =  infoFactura.find('totalSinImpuestos').text
                        else:
                            totalSinImpuestos = '0'
                        worksheet.write('T'+str(count), totalSinImpuestos, data_format)

                        if infoFactura.find('propina') is not None:
                            propina = infoFactura.find('propina').text
                        else:
                            propina = '0.00'
                        worksheet.write('Y'+str(count), propina, data_format)

                        if infoFactura.find('importeTotal') is not None:
                            importeTotal = infoFactura.find('importeTotal').text
                        else:
                            importeTotal = '0.00'
                        worksheet.write('Z'+str(count), importeTotal, data_format)

                        if infoFactura.find('totalSubsidio') is not None:
                            totalSubsidio = infoFactura.find('totalSubsidio').text
                        else:
                            totalSubsidio = '0.00'
                        worksheet.write('AB'+str(count), totalSubsidio, data_format)

                        valorSinSubsidio = float(importeTotal) - float(totalSubsidio)
                        worksheet.write('AA'+str(count), valorSinSubsidio, data_format)

                        if infoFactura.find('totalConImpuestos') is not None:
                            countIce = 0
                            countIRBPNR = 0
                            for totalImpuesto in infoFactura.find('totalConImpuestos').findall('totalImpuesto'):
                                if totalImpuesto.find('codigo') is not None:
                                    
                                    if int(totalImpuesto.find('codigo').text) == 2: 
                                        
                                        # iva 12%
                                        if int(totalImpuesto.find('codigoPorcentaje').text) == 2:    
                                            iva =totalImpuesto.find('valor').text
                                            baseImponible = totalImpuesto.find('baseImponible').text
                                            subtotal_doce = float(iva) + float(baseImponible)
                                            worksheet.write('P'+str(count), subtotal_doce, data_format)
                                            worksheet.write('W'+str(count), iva, data_format)
                                        else:
                                            worksheet.write('P'+str(count), '0.00', data_format)
                                            worksheet.write('W'+str(count), '0.00', data_format)
                                        
                                        # iva 0%
                                        if int(totalImpuesto.find('codigoPorcentaje').text) == 0:    
                                            subtotal_cero = totalImpuesto.find('baseImponible').text
                                            worksheet.write('Q'+str(count), subtotal_cero, data_format)
                                        else:
                                            worksheet.write('Q'+str(count), '0.00', data_format)

                                        # no objeto de iva
                                        if int(totalImpuesto.find('codigoPorcentaje').text) == 6:
                                            noObjetoIva = totalImpuesto.find('baseImponible').text
                                            worksheet.write('R'+str(count), noObjetoIva, data_format)
                                        else:
                                            worksheet.write('R'+str(count), '0.00', data_format)
                                        
                                        # exento de iva
                                        if int(totalImpuesto.find('codigoPorcentaje').text) == 7:
                                            exentoIva = totalImpuesto.find('baseImponible').text
                                            worksheet.write('S'+str(count), exentoIva, data_format)
                                        else:
                                            worksheet.write('S'+str(count), '0.00', data_format)

                                    if (int(totalImpuesto.find('codigo').text) == 3 ):
                                        ice = totalImpuesto.find('valor').text
                                        worksheet.write('V'+str(count), ice, data_format)
                                        countIce +=1
                                    
                                    if int(totalImpuesto.find('codigo').text) == 5:
                                        IRBPNR = totalImpuesto.find('valor').text
                                        worksheet.write('X'+str(count), IRBPNR, data_format)
                                        countIRBPNR += 1
                            
                            if countIce == 0:
                                worksheet.write('V'+str(count), '0.00', data_format)
                            
                            if countIRBPNR == 0:
                                worksheet.write('X'+str(count), '0.00', data_format)

                    
                    if infoAdicional is not None:
                        adicional = ''
                        for campoAdicional in infoAdicional.findall('campoAdicional'):
                            adicional += campoAdicional.get('nombre') + ':  ' + campoAdicional.text + '\n'
                    else:
                        adicional = ''
                    worksheet.write('AC'+str(count), adicional, data_format)

            workbook.close()

            return HttpResponse(1)
        
        else:
            return HttpResponse(2)
    else:
        return HttpResponse(0)

def bodyHeader(arg):
    
    
    productsLeftStyle = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 7, alignment = TA_LEFT )
    productsCenterStyle = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 7, alignment = TA_CENTER, )
    productosRightStyle = ParagraphStyle('parrafo', fontName = "Helvetica", fontSize = 7, alignment = TA_RIGHT, )

    datos = []

    claveAcceso = ''
    if arg[0] == 'Factura':
        claveAcceso = arg[8]
    if arg[0] == 'Comprobante de Retención':
        claveAcceso = arg[9]
    if arg[0] == 'Notas de Crédito':
        claveAcceso = arg[9]
    if arg[0] == 'Notas de Débito':
        claveAcceso = arg[9]

    
    headers = {'Content-Type': 'application/xml','Accept': 'application/xml'}
    body = "<Envelope xmlns=\"http://schemas.xmlsoap.org/soap/envelope/\">"
    body += "    <Body>"
    body += "       <autorizacionComprobante xmlns=\"http://ec.gob.sri.ws.autorizacion\">"
    body += "           <claveAccesoComprobante xmlns=\"\">"+claveAcceso+"</claveAccesoComprobante>"
    body += "       </autorizacionComprobante>"
    body += "    </Body>"
    body += "</Envelope>"
    r = requests.post(url="https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl", data=body, headers=headers)
    xml_response = r.text

    ns = {'soap':'http://schemas.xmlsoap.org/soap/envelope/'}
    ns2 = {'ns2': 'http://ec.gob.sri.ws.autorizacion'}
    root = ElementTree.fromstring(xml_response)

    numeroAutorizacion = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('numeroAutorizacion').text
    datos.append(numeroAutorizacion)

    fechaAutorizacion = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('fechaAutorizacion').text
    datos.append(fechaAutorizacion)

    ambiente = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('ambiente').text
    datos.append(ambiente)

    claveAccesoConsultada = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('claveAccesoConsultada').text
    
    value = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('comprobante')
    
    if value is not None:

        text2 = value.text.replace("&lt;","<")
        rootFactura = ElementTree.fromstring(text2)
        detalles = rootFactura.find('detalles')
        infoTributaria = rootFactura.find('infoTributaria')
        infoFactura = rootFactura.find('infoFactura')
        infoAdicional = rootFactura.find('infoAdicional')

        if infoTributaria is not None:

            if infoTributaria.find('razonSocial') is not None:
                razonSocial = infoTributaria.find('razonSocial').text
            else: 
                razonSocial = ''  
            datos.append(razonSocial) 

            if infoTributaria.find('dirMatriz') is not None:
                dirMatriz = infoTributaria.find('dirMatriz').text
            else:
                dirMatriz = ''
            datos.append(dirMatriz)

            if infoTributaria.find('tipoEmision') is not None:
                emisionNumero = int(infoTributaria.find('tipoEmision').text)
                if emisionNumero == 1:
                    tipoEmision = 'NORMAL'
                else:
                    tipoEmision = ''
            else:
                tipoEmision = '' 
            datos.append(tipoEmision)                    

            if infoTributaria.find('ruc') is not None:
                ruc = infoTributaria.find('ruc').text
            else:
                ruc = ''
            datos.append(ruc) 
            
            if infoTributaria.find('codDoc') is not None:
                codDoc = int(infoTributaria.find('codDoc').text)
            else:
                codDoc = 0
            datos.append(codDoc) 
                
            if codDoc == 1:
                codigDoc = 'FACTURA'
            if codDoc == 4:
                codigDoc = 'NOTA DE CRÉDITO'
            if codDoc == 5:
                codigDoc = 'NOTA DE DÉBITO'
            if codDoc == 6:
                codigDoc = 'GUÍA DE REMISIÓN'
            if codDoc == 7:
                codigDoc = 'COMPROBANTE DE RETENCIÓN'
            if codDoc == 0:
                codigDoc = ''
            datos.append(codigDoc)

            if (infoTributaria.find('estab') is not None and
                infoTributaria.find('ptoEmi') is not None and
                infoTributaria.find('secuencial') is not None):
                estab = infoTributaria.find('estab').text
                ptoEmi = infoTributaria.find('ptoEmi').text
                secuencial = infoTributaria.find('secuencial').text
                No = estab + '-' + ptoEmi + '-' + secuencial 
            else:
                No = ''  
            datos.append(No)

            if infoTributaria.find('nombreComercial') is not None:
                nombreComercial = infoTributaria.find('nombreComercial').text
            else:
                nombreComercial = ''
            datos.append(nombreComercial)
        
        if infoFactura is not None:

            if infoFactura.find('dirEstablecimiento') is not None:
                establecimiento = infoFactura.find('dirEstablecimiento').text
                if establecimiento != dirMatriz:
                    dirEstablecimiento = infoFactura.find('dirEstablecimiento').text
                else:
                    dirEstablecimiento = ''
            else:
                dirEstablecimiento = ''
            datos.append(dirEstablecimiento)

            if infoFactura.find('direccionComprador') is not None: 
                direccionComprador = infoFactura.find('direccionComprador').text
            else:
                direccionComprador = ''
            datos.append(direccionComprador)

            if infoFactura.find('obligadoContabilidad') is not None: 
                obligadoContabilidad = infoFactura.find('obligadoContabilidad').text
            else:
                obligadoContabilidad = ''
            datos.append(obligadoContabilidad)

            if infoFactura.find('razonSocialComprador') is not None:
                razonSocialComprador = infoFactura.find('razonSocialComprador').text
            else:
                razonSocialComprador = ''
            datos.append(razonSocialComprador)

            if infoFactura.find('identificacionComprador') is not None: 
                identificacionComprador = infoFactura.find('identificacionComprador').text
            else:
                identificacionComprador = ''
            datos.append(identificacionComprador)

            if infoFactura.find('fechaEmision') is not None:
                fechaEmision = infoFactura.find('fechaEmision').text
            else:
                fechaEmision = ''
            datos.append(fechaEmision)

            if infoFactura.find('totalDescuento') is not None:
                totalDescuento = "{0:.2f}".format(float(infoFactura.find('totalDescuento').text))
            else:
                totalDescuento = '0'
            datos.append(totalDescuento)

            if infoFactura.find('pagos').find('pago').find('formaPago') is not None:
                formaPagoNumero = int(infoFactura.find('pagos').find('pago').find('formaPago').text)
                if formaPagoNumero == 1:
                    formaPago = 'SIN UTILIZACION DEL SISTEMA FINANCIERO'
                elif formaPagoNumero == 15:
                    formaPago = 'COMPENSACIÓN DE DEUDAS'
                elif formaPagoNumero == 16:
                    formaPago = 'TARJETA DE DÉBITO'
                elif formaPagoNumero == 17:
                    formaPago = 'DINERO ELECTRÓNICO'
                elif formaPagoNumero == 18:
                    formaPago = 'TARJETA PREPAGO'
                elif formaPagoNumero == 19:
                    formaPago = 'TARJETA DE CRÉDITO'
                elif formaPagoNumero == 20:
                    formaPago = 'OTROS CON UTILIZACION DEL SISTEMA FINANCIERO'
                elif formaPagoNumero == 21:
                    formaPago = 'ENDOSO DE TÍTULOS'
                else:
                    formaPago = ''
            else:
                formaPago = ''
            datos.append(formaPago)

            if infoFactura.find('pagos').find('pago').find('total') is not None:
                total0 = float(infoFactura.find('pagos').find('pago').find('total').text)
                total ="{0:.2f}".format(total0)
            else:
                total = '0'
            datos.append(total)

            if infoFactura.find('totalSinImpuestos') is not None:
                    totalSinImpuestos = "{0:.2f}".format(float(infoFactura.find('totalSinImpuestos').text))
            else:
                totalSinImpuestos = '0'
            datos.append(totalSinImpuestos)

            if infoFactura.find('propina') is not None:
                propina = "{0:.2f}".format(float(infoFactura.find('propina').text))
            else:
                propina = '0.00'
            datos.append(propina)

            if infoFactura.find('importeTotal') is not None:
                importeTotal = "{0:.2f}".format(float(infoFactura.find('importeTotal').text))
            else:
                importeTotal = '0.00'
            datos.append(importeTotal)

            if infoFactura.find('totalSubsidio') is not None:
                totalSubsidio = "{0:.2f}".format(float(infoFactura.find('totalSubsidio').text))
            else:
                totalSubsidio = '0.00'
            datos.append(totalSubsidio)

            valorSinSubsidio = float(importeTotal) - float(totalSubsidio)
            datos.append(valorSinSubsidio)

            if infoFactura.find('totalConImpuestos') is not None:
                countIce = 0
                countIRBPNR = 0
                arregloImpuesto = []
                for totalImpuesto in infoFactura.find('totalConImpuestos').findall('totalImpuesto'):
                    if totalImpuesto.find('codigo') is not None:
                        
                        if int(totalImpuesto.find('codigo').text) == 2: 
                            
                            # iva 12%
                            if int(totalImpuesto.find('codigoPorcentaje').text) == 2:    
                                iva = "{0:.2f}".format(float(totalImpuesto.find('valor').text))
                                baseImponible = totalImpuesto.find('baseImponible').text
                                subtotal_doce = "{0:.2f}".format(float(iva) + float(baseImponible))
                            else:
                                iva = '0.00'
                                subtotal_doce = '0.00'
                            tuplaIva = ("iva", iva)
                            tuplaSubDoce = ("subDoce", subtotal_doce)
                            arregloImpuesto.append(tuplaIva)
                            arregloImpuesto.append(tuplaSubDoce)
                            
                            # iva 0%
                            if int(totalImpuesto.find('codigoPorcentaje').text) == 0:    
                                subtotal_cero = "{0:.2f}".format(float(totalImpuesto.find('baseImponible').text))
                            else:
                                subtotal_cero = '0.00'
                            tuplaSubCero = ("subCero", subtotal_cero)
                            arregloImpuesto.append(tuplaSubCero)

                            # no objeto de iva
                            if int(totalImpuesto.find('codigoPorcentaje').text) == 6:
                                noObjetoIva = "{0:.2f}".format(float(totalImpuesto.find('baseImponible').text))
                            else:
                                noObjetoIva = '0.00'
                            tuplaNoIva = ("noIva", noObjetoIva)
                            arregloImpuesto.append(tuplaNoIva)
                            
                            # exento de iva
                            if int(totalImpuesto.find('codigoPorcentaje').text) == 7:
                                exentoIva = totalImpuesto.find('baseImponible').text
                            else:
                                exentoIva = '0.00'
                            tuplaExeIva = ("exeIva", exentoIva)
                            arregloImpuesto.append(tuplaExeIva)

                        if (int(totalImpuesto.find('codigo').text) == 3 ):
                            ice = "{0:.2f}".format(float(totalImpuesto.find('valor').text))
                            countIce +=1
                        
                        if int(totalImpuesto.find('codigo').text) == 5:
                            IRBPNR = totalImpuesto.find('valor').text
                            countIRBPNR += 1
                
                datos.append(arregloImpuesto)
                
                if countIce == 0:
                    ice = '0.00'
                datos.append(ice)
                
                if countIRBPNR == 0:
                    IRBPNR = '0.00'
                datos.append(IRBPNR)

        if detalles is not None:
            
            detalleArray = []
            contadorDetalles = 0
            
            for child in detalles.findall('detalle'):
                contadorDetalles += 1
                detalleAd = ''
                detallesArray = []

                if child.find('codigoPrincipal') is not None:
                    codigoPrincipal = child.find('codigoPrincipal').text
                else:
                    codigoPrincipal = ''
                detalle1 = Paragraph(codigoPrincipal, productsCenterStyle)

                if child.find('codigoAuxiliar') is not None:
                    codigoAuxiliar = child.find('codigoAuxiliar').text
                else:
                    codigoAuxiliar = ''
                detalle2 = Paragraph(codigoAuxiliar, productsCenterStyle)

                if child.find('cantidad') is not None:
                    can = float(child.find('cantidad').text)
                    cantidad = "{0:.2f}".format(can)
                else:
                    cantidad = '0'
                detalle3 = Paragraph(cantidad, productsCenterStyle)

                if child.find('descripcion') is not None:
                    descripcion = child.find('descripcion').text
                else:
                    descripcion = ''
                detalle4 = Paragraph(descripcion, productsLeftStyle)

                if child.find('detallesAdicionales') is not None:
                    for child2 in child.find('detallesAdicionales').findall('detAdicional'):
                        if child2 is not None:
                            nombre = child2.get('nombre')
                            valor = child2.get('valor=')
                            detalleAd += str(nombre) + ':   '+ str(valor) +'\n'
                
                detalle5 = Paragraph(detalleAd, productsLeftStyle)

                if child.find('precioUnitario') is not None:
                    precUni = float(child.find('precioUnitario').text)
                    precioUnitario = "{0:.2f}".format(precUni)
                else:
                    precioUnitario = '0.00'
                detalle6 = Paragraph(precioUnitario, productosRightStyle) 

                if child.find('precioSinSubsidio') is not None:
                    precioSinSub = float(child.find('precioSinSubsidio').text)
                    precioSinSubsidio = "{0:.2f}".format(precioSinSub)
                else:
                    precioSinSubsidio = '0.00' 
                detalle8 = Paragraph(precioSinSubsidio, productosRightStyle)

                if float(precioSinSubsidio) == 0.00:
                    subsidio = '0.00'
                else:
                    subsidio = "{0:.2f}".format(float(precioSinSubsidio) - float(precioTotalSinImpuesto))
                detalle7 = Paragraph(subsidio, productosRightStyle)                         
                
                if child.find('descuento') is not None:
                    descue = float(child.find('descuento').text)
                    descuento = "{0:.2f}".format(descue)
                else:
                    descuento = ''
                detalle9 = Paragraph(descuento, productosRightStyle)

                if child.find('precioTotalSinImpuesto') is not None:
                    precioTotalSinIm = float(child.find('precioTotalSinImpuesto').text)
                    precioTotalSinImpuesto = "{0:.2f}".format(precioTotalSinIm)
                else:
                    precioTotalSinImpuesto = '0.00'
                detalle10 = Paragraph(precioTotalSinImpuesto, productosRightStyle)

                detallesArray.append(detalle1)
                detallesArray.append(detalle2)
                detallesArray.append(detalle3)
                detallesArray.append(detalle4)
                detallesArray.append(detalle5)
                detallesArray.append(detalle6)
                detallesArray.append(detalle7)
                detallesArray.append(detalle8)
                detallesArray.append(detalle9)
                detallesArray.append(detalle10)
                detalleArray.append(detallesArray)
        
        datos.append(detalleArray)
        
        if infoAdicional is not None:
            adicional = ''
            for campoAdicional in infoAdicional.findall('campoAdicional'):
                adicional += campoAdicional.get('nombre') + ':  ' + campoAdicional.text + '<br />\n'
        else:
            adicional = ''
        datos.append(adicional)
    
    
    return (datos)

def bodyHeader2(arg):
    headers = {'Content-Type': 'application/xml','Accept': 'application/xml'}
    body = "<Envelope xmlns=\"http://schemas.xmlsoap.org/soap/envelope/\">"
    body += "    <Body>"
    body += "       <autorizacionComprobante xmlns=\"http://ec.gob.sri.ws.autorizacion\">"
    body += "           <claveAccesoComprobante xmlns=\"\">"+arg+"</claveAccesoComprobante>"
    body += "       </autorizacionComprobante>"
    body += "    </Body>"
    body += "</Envelope>"
    r = requests.post(url="https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline?wsdl", data=body, headers=headers)
    return (r.text)
