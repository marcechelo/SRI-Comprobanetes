from django.urls import path
from . import views
from django.conf.urls import url
from sri_test.views import sri_document
from django.views.generic import RedirectView
from django.urls import re_path

app_name = 'sri_test'

urlpatterns = [
    path('', sri_document.as_view(), name='index'),
    path('comprobantes/', views.test, name= 'comprobantes'),
    path('downloadxml/', views.downloadxml, name= 'downloadxml'),
    path('downloadPdf/', views.downloadPdf, name= 'downloadPdf'),
    path('downloadeExcel/', views.downloadeExcel, name = 'downloadeExcel'),
    path('emitidos/', views.comprobantesEmitidos, name= 'emitidos'),
    path('recibidos/', views.comprobRecibido, name= 'recibidos'),

    path('read/', views.read_txt, name= 'read'),
    path('pdf/', views.getpdf, name='pdf'),

]