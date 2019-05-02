from django.urls import path
from . import views
from django.conf.urls import url
from sri_test.views import sri_document

app_name = 'sri_test'

urlpatterns = [
    path('', views.index, name='index'),
    path('test/', views.test, name= 'test'),
    path('reload/', views.reload, name= 'reload'),
    path('downloadxml/', views.downloadxml, name= 'downloadxml'),
    path('downloadPdf/', views.downloadPdf, name= 'downloadPdf'),
    path('emitidos/', views.comprobantesEmitidos, name= 'emitidos'),
    path('recibidos/', views.comprobantesRecibidos, name= 'recibidos'),

    path('read/', views.read_txt, name= 'read'),
    path('createPDF/', views.createPDF, name = 'createPDF'),
    path('createExcel/', views.createExcel, name = 'createExcel'),

    path('path/', views.selectPath, name = 'path'),
]