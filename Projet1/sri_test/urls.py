from django.urls import path
from . import views
from django.conf.urls import url
from sri_test.views import sri_document

app_name = 'sri_test'

urlpatterns = [
    path('', views.index, name='index'),
    path('comprobantesRecibidos/', views.test, name= 'comprobantesRecibidos'),
    path('downloadxml/', views.downloadxml, name= 'downloadxml'),
    path('downloadPdf/', views.downloadPdf, name= 'downloadPdf'),
    path('downloadeExcel/', views.downloadeExcel, name = 'downloadeExcel'),
    path('emitidos/', views.comprobantesEmitidos, name= 'emitidos'),
    path('recibidos/', views.comprobRecibido, name= 'recibidos'),

    path('read/', views.read_txt, name= 'read'),

]