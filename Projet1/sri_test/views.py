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
from reportlab.lib.enums import TA_CENTER

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


import os, os.path
import requests
import xml.dom.minidom
import json
from xml.etree import ElementTree
import xlsxwriter
import reportlab


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
#
# def make_pdf(xml_file,auth_number):
#     """
#     Create a PDF based on the XML data
#     """
#     self_canvas = canvas.Canvas(auth_number+".pdf", pagesize=letter)
#     width, height = letter
#     styles = getSampleStyleSheet()
#     xml_bytes = bytes(bytearray(xml_file, encoding = 'utf-8'))
#     xml_object = objectify.fromstring(xml_bytes)
#     address = """ <font size="9">
#     SHIP TO:<br/>
#     <br/>
#     %s<br/>
#     %s<br/>
#     %s<br/>
#     %s<br/>
#     </font>
#     """ % (xml_object.ambiente, xml_object.ambiente, xml_object.ambiente, xml_object.ambiente)
#     p = Paragraph(address, styles["Normal"])
#     p.wrapOn(self_canvas, width, height)
#     p.drawOn(self_canvas, *coord(18, 40, mm,height))
#
#     order_number = '<font size="14"><b>Order #%s </b></font>' % xml_object.ambiente
#     p = Paragraph(order_number, styles["Normal"])
#     p.wrapOn(self_canvas, width, height)
#     p.drawOn(self_canvas, * coord(18, 50, mm,height))
#
#     data = []
#     data.append(["Item ID", "Name", "Price", "Quantity", "Total"])
#     grand_total = 0
#     for item in xml_object.detalles.iterchildren():
#         row = []
#         row.append(item.codigoPrincipal)
#         row.append(item.descripcion)
#         row.append(item.cantidad)
#         row.append(item.precioUnitario)
#         total = Decimal(str(item.precioUnitario)) * Decimal(str(item.cantidad))
#         row.append(str(total))
#         grand_total += total
#         data.append(row)
#     data.append(["", "", "", "Grand Total:", grand_total])
#     t = Table(data, 1.5 * inch)
#     t.setStyle(TableStyle([
#         ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
#         ('BOX', (0,0), (-1,-1), 0.25, colors.black)
#     ]))
#     t.wrapOn(self_canvas, width, height)
#     t.drawOn(self_canvas, *coord(18, 85, mm,height))
#
#     txt = "Thank you for your business!"
#     p = Paragraph(txt, styles["Normal"])
#     p.wrapOn(self_canvas, width, height)
#     p.drawOn(self_canvas, *coord(18, 95, mm,height))
#
# def coord(self, x, y, unit=1,height):
#     """
#     # http://stackoverflow.com/questions/4726011/wrap-text-in-a-table-reportlab
#     Helper class to help position flowables in Canvas objects
#     """
#     x, y = x * unit, height -  y * unit
#     return x, y


#Read from selected file

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

def get_download_path(query):
    print("Hello World!")
    result = []
    return HttpResponse("done")

def reload(request):
    return render(request, 'sri_test/index.html')

def downloadxml(request):
    global fileUploaded, dataDocumentArray
    if request.method == 'GET':
        for i in dataDocumentArray:
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
            xml_response = xml_response.replace("&#xd;","")
            with open(os.path.join(os.path.join(os.path.expanduser('~'),'Downloads',i[8]+".xml")), "w+") as file1:
                file1.write(xml_response)
            #f = open(i[8]+".xml","w+")
            #f.write(xml_response)
        #print(dataDocumentArray)
        context = {'message': 'ok'}
        return HttpResponse(context)

def downloadPdf(request):
    global fileUploaded, dataDocumentArray
    count = 0
    countruc = 0
    if request.method == 'GET':
        for i in dataDocumentArray[1:]:
            count += 1
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
            #xml_response = xml_response.replace

            ns = {'soap':'http://schemas.xmlsoap.org/soap/envelope/'}
            ns2 = {'ns2': 'http://ec.gob.sri.ws.autorizacion'}
            root = ElementTree.fromstring(xml_response)

            numeroAutorizacion = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('numeroAutorizacion').text
            fechaAutorizacion = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('fechaAutorizacion').text
            ambiente = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('ambiente').text
            claveAccesoConsultada = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('claveAccesoConsultada').text
            print(str(count) + '  '+ claveAccesoConsultada)
            value = root.find('soap:Body', ns).find('{http://ec.gob.sri.ws.autorizacion}autorizacionComprobanteResponse').find('RespuestaAutorizacionComprobante').find('autorizaciones').find('autorizacion').find('comprobante')
            
            if value is not None:

                text2 = value.text.replace("&lt;","<")
                rootFactura = ElementTree.fromstring(text2)
                detalles = rootFactura.find('detalles')
                infoTributaria = rootFactura.find('infoTributaria')
                infoFactura = rootFactura.find('infoFactura')
                infoAdicional = rootFactura.find('infoAdicional')

                if infoTributaria is not None:

                    if infoTributaria.find('ruc') is not None:
                        ruc = infoTributaria.find('ruc').text
                    else:
                        ruc = ''

                    if infoTributaria.find('razonSocial') is not None:
                        razonSocial = infoTributaria.find('razonSocial').text
                    else: 
                        razonSocial = ''
                    
                    if infoTributaria.find('dirMatriz') is not None:
                        dirMatriz = infoTributaria.find('dirMatriz').text
                    else:
                        dirMatriz = ''

                    if infoTributaria.find('tipoEmision') is not None:
                        tipoEmision = infoTributaria.find('tipoEmision').text
                    else:
                        tipoEmision = ''

                    if (infoTributaria.find('estab') is not None and
                        infoTributaria.find('ptoEmi') is not None and
                        infoTributaria.find('secuencial') is not None):
                        estab = infoTributaria.find('estab').text
                        ptoEmi = infoTributaria.find('ptoEmi').text
                        secuencial = infoTributaria.find('secuencial').text
                        No = estab + '-' + ptoEmi + '-' + secuencial 
                    else:
                        No = ''     

                    if infoTributaria.find('nombreComercial') is not None:
                        nombreComercial = infoTributaria.find('nombreComercial').text
                    else:
                        nombreComercial = ''

                    if infoTributaria.find('ambiente') is not None:
                        ambiente = infoTributaria.find('ambiente').text
                    else:
                        ambiente = ''
                
                if infoFactura is not None:
                      
                    if infoFactura.find('direccionComprador') is not None: 
                        direccionComprador = infoFactura.find('direccionComprador').text
                    else:
                        direccionComprador = ''

                    if infoFactura.find('obligadoContabilidad') is not None: 
                        obligadoContabilidad = infoFactura.find('obligadoContabilidad').text
                    else:
                        obligadoContabilidad = ''

                    if infoFactura.find('razonSocialComprador') is not None:
                        razonSocialComprador = infoFactura.find('razonSocialComprador').text
                    else:
                        razonSocialComprador = ''

                    if infoFactura.find('identificacionComprador') is not None: 
                        identificacionComprador = infoFactura.find('identificacionComprador').text
                    else:
                        identificacionComprador = ''

                    if infoFactura.find('fechaEmision') is not None:
                        fechaEmision = infoFactura.find('fechaEmision').text
                    else:
                        fechaEmision = ''

                    if infoFactura.find('totalDescuento') is not None:
                        totalDescuento = infoFactura.find('totalDescuento').text
                    else:
                         totalDescuento = '0'

                    if infoFactura.find('pagos').find('pago').find('formaPago') is not None:
                        formaPago = infoFactura.find('pagos').find('pago').find('formaPago').text
                    else:
                        formaPago = '' 

                    if infoFactura.find('pagos').find('pago').find('total') is not None:
                        total = infoFactura.find('pagos').find('pago').find('total').text
                    else:
                        total = '0' 
                    if infoFactura.find('totalSinImpuestos') is not None:
                         totalSinImpuestos =  infoFactura.find('totalSinImpuestos').text
                    else:
                        totalSinImpuestos = '0'

                    if infoFactura.find('propina') is not None:
                        propina = infoFactura.find('propina').text
                    else:
                        propina = '0.00'

                    if infoFactura.find('importeTotal') is not None:
                        importeTotal = infoFactura.find('importeTotal').text
                    else:
                        importeTotal = '0.00'

                    if infoFactura.find('totalConImpuestos') is not None:
                        countruc += 1
                        for totalImpuesto in infoFactura.find('totalConImpuestos').findall('totalImpuesto'):
                            if totalImpuesto.find('codigo') is not None:
                                exentoIva = ''
                                if (int(totalImpuesto.find('codigo').text) == 2 and 
                                    float(totalImpuesto.find('baseImponible').text) != 0 and
                                    float(totalImpuesto.find('valor').text) != 0):
                                    #totalImpuesto.find('valor') is not None):
                                    iva =totalImpuesto.find('valor').text
                                    subtotal_cero = totalImpuesto.find('baseImponible').text
                                    subtotal_doce = float(totalImpuesto.find('baseImponible').text) + float(totalImpuesto.find('valor').text)
                                else:
                                    iva = 'not'
                                    subtotal_cero = '0.00'
                                    subtotal_doce = 0.00
                                #print(str(countruc)+ '  iva: '+ iva + '  subtotal 0: '+ subtotal_cero + '  subtotal 12: '+ str(subtotal_doce))
                                
                                if (int(totalImpuesto.find('codigo').text) == 2 and 
                                    float(totalImpuesto.find('baseImponible').text) != 0 and
                                    float(totalImpuesto.find('valor').text) == 0):
                                    exentoIva += ' ' + totalImpuesto.find('baseImponible').text
                                    countruc += 1
                                    print(str(count) + ' '+ exentoIva)
                                

                                if (int(totalImpuesto.find('codigo').text) == 3 and
                                    float(totalImpuesto.find('baseImponible').text) != 0):
                                    #totalImpuesto.find('valor') is not None):
                                    ice = totalImpuesto.find('valor').text
                                else:
                                    ice = 'not'
                            
                if detalles is not None:
                    for child in detalles.findall('detalle'):
                        detalleAd = ''

                        if child.find('detallesAdicionales') is not None:
                            for child2 in child.find('detallesAdicionales'):
                                if child2.find('detAdicional') is not None:
                                    detAdicional = child2.find('detAdicional').text
                                    detalleAd += detAdicional + '\n'
                                
                        #print(detalleAd)

                        if (child.find('codigoPrincipal') is not None and 
                            child.find('cantidad') is not None and
                            child.find('precioUnitario') is not None and
                            child.find('descripcion') is not None):
                            
                            precioUnitario = child.find('precioUnitario').text
                            cantidad = child.find('cantidad').text
                            codigoPrincipal = child.find('codigoPrincipal').text
                            descripcion = child.find('descripcion').text
                        
                        #print(descripcion)
                
                if infoAdicional is not None:
                    adicional = ''
                    for campoAdicional in infoAdicional.findall('campoAdicional'):
                        adicional += campoAdicional.get('nombre') + ':  ' + campoAdicional.text + '\n'
            
            #with open(os.path.join(os.path.join(os.path.expanduser('~'),'Downloads',i[8]+".xml")), "w+") as file1:
            #    file1.write(xml_response)
            #f = open(i[8]+".xml","w+")
            #f.write(xml_response)
        #print(dataDocumentArray)
        context = {'message': 'ok'}
        return HttpResponse(context)

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

def createPDF(request):

    c = canvas.Canvas("Factura.pdf", pagesize=A4)
    c.translate(0,(0.7)*inch)
    #c.rotate(180)

    #rect (w,h)
    c.roundRect(4*inch, (6.5)*inch, 287, 300, 4)
    c.roundRect((0.2)*inch, (6.5)*inch, 260, 200, 4)
    c.rect((0.2)*inch, (5.4)*inch, 560, 75)

    #Paragraph style
    style1 = ParagraphStyle('parrafo', fontName = "Helvetica-Bold", fontSize = 10 )
    style2 = ParagraphStyle('parrafo', fontName = "Helvetica-Bold", fontSize = 8 )
    style3 = ParagraphStyle('parrafo', fontName = "Helvetica-Bold", fontSize = 7, alignment = TA_CENTER, )

    c.setFont("Helvetica", 8)
    c.setFillColorRGB(255,0,0)
    
    #Logo
    c.setFont("Helvetica-Bold", 28)
    c.setFillColorRGB(255,0,0)
    message = 'NO TIENE LOGO'
    c.drawString(0.5*inch, (10.1)*inch, message)

    #First Square

    message = '####################################################################'
    #styles = getSampleStyleSheet()    
    p = Paragraph(message, style1)
    p.wrapOn(c, (3.4)*inch, (2.5)*inch)  # size of 'textbox' for linebreaks etc.
    p.drawOn(c, (0.3)*inch, (8.9)*inch)

    p = Paragraph(message, style2)
    p.wrapOn(c, (3.4)*inch, (2.5)*inch)
    p.drawOn(c, (0.3)*inch, (8.3)*inch)

    c.setFont("Helvetica", 8)
    c.setFillColorRGB(0,0,0)
    c.drawString((0.3)*inch, (8.0)*inch, 'Dirección')
    c.drawString((0.3)*inch, (7.9)*inch, 'Matriz:')

    message = 'Here goes the direction ############################################################################################################'
    p = Paragraph(message, style2)
    p.wrapOn(c, (2.5)*inch, (2.5)*inch)
    p.drawOn(c, (1.2)*inch, (7.5)*inch)
    
    c.drawString((0.3)*inch, (7.3)*inch, 'Dirección')
    c.drawString((0.3)*inch, (7.2)*inch, 'Sucursal:')

    message = 'Here goes the direction 2############################################################################################################'
    p = Paragraph(message, style2)
    p.wrapOn(c, (2.5)*inch, (2.5)*inch)
    p.drawOn(c, (1.2)*inch, (6.8)*inch)

    message = 'OBLIGADO A LLEVAR:                   ' + 'here goes yes/no'
    c.drawString((0.3)*inch, (6.6)*inch, message)
    
    #Second Square

    c.setFont("Helvetica", 14)
    c.setFillColorRGB(0,0,0)
    message = 'R.U.C.: ' +'Here egoes the ruc'
    c.drawString((4.1)*inch, (10.3)*inch, message)
    c.drawString((4.1)*inch, (10)*inch, 'FACTURA')

    c.setFont("Helvetica", 10)
    message = 'No.  ' + 'Here goes the No'
    c.drawString((4.1)*inch, (9.7)*inch, message)

    message = 'NÚMERO DE AUTORIZACIÓN'
    c.drawString((4.1)*inch, (9.4)*inch, message)

    message = 'here goes autorisation number'
    c.drawString((4.1)*inch, (9.1)*inch, message)

    message = 'FECHA Y HORA DE'
    c.drawString((4.1)*inch, (8.8)*inch, message)

    message = 'AUTORIZACIÓN:               ' + 'here goes autorisation date'
    c.drawString((4.1)*inch, (8.6)*inch, message)

    message = 'AMBIENTE:                        ' + 'here goes the enviroment'
    c.drawString((4.1)*inch, (8.3)*inch, message)

    message = 'EMISIÓN:                           ' + 'here goes the emision'
    c.drawString((4.1)*inch, (8.0)*inch, message)

    message = 'CLAVE DE ACCESO'
    c.drawString((4.1)*inch, (7.7)*inch, message)

    message = 'HERE GOES THE IMAGE'
    c.drawString((4.1)*inch, (6.9)*inch, message)

    message = 'here goes access password'
    c.drawString((4.1)*inch, (6.7)*inch, message)

    c.translate(0*inch, 0*inch)

    #Third square
    
    c.setFont("Helvetica", 8)
    c.setFillColorRGB(0,0,0)
    message = 'Razón Social/Nombres                                  '+'here goes the name'
    c.drawString((0.3)*inch, (6.3)*inch, message)

    message = 'Identificacion                 '+'here goes the id'
    c.drawString((0.3)*inch, (6.1)*inch, message)

    message = 'Fecha                            '+'here goes the date'+'                           Placa/Matrícula:                        '+'here goes lisence plate'
    c.drawString((0.3)*inch, (5.9)*inch, message)
 
    message = 'Dirección                       '+'here goes the direction'
    c.drawString((0.3)*inch, (5.7)*inch, message)

    #products square width = 560
    
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
    
    size = A4
    print (A4[0])
    print(size[1]*mm)


    data = [[p, p1, p2, p3, p4, p5, p6,p7,p8, p9], [2, 1, 3,4,5,6,7,8,9,0], [3, 2, 1,4,5,6,7,8,9]]

    table = Table(data, colWidths=[50, 50, 50, 80, 80, 50, 50, 50, 50, 50, 50])
    table.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                    ("ALIGN", (0,0), (-1,-1), "CENTER"),
                    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                    ('BOX', (0,0), (-1,-1), 0.25, colors.black)])

    table.wrapOn(c, size[0], size[1])
    table.drawOn(c, (0.2)*inch, (4.5)*inch)

    table.canv = c
    w, h = table.wrap(0,0)

    print(h)
    
    c.save()

    return HttpResponse('DONE')

def createExcel(request):

    workbook = xlsxwriter.Workbook('factura.xlsx') 
    worksheet = workbook.add_worksheet() 
    
    #Formats
    data_format = workbook.add_format({'text_wrap': True, 'font_size': 8, 'border': 1, 'valign': 'vcenter'})
    titles_format = workbook.add_format({'text_wrap': True, 'bold': 1, 'align': 'center', 'valign': 'vcenter','font_size': 10, 'border': 1,})
    merge_format = workbook.add_format({'text_wrap': True, 'bold': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 12, 'border': 1,})
    
    #Tamaños de las columnas
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
    worksheet.write('A2', '#', titles_format) 
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
        worksheet.write('A'+str(count), count-2, data_format) 
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
        #xml_response = xml_response.replace

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
                                    subtotal_doce = totalImpuesto.find('baseImponible').text
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

    return HttpResponse('DONE')

def selectPath(self):
    root = tkinter.Tk()
    dirname = filedialog.askdirectory(parent=root, initialdir="/", title='Please select a directory')
    print(dirname)
    return HttpResponse('DONE')