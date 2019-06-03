{\rtf1\ansi\ansicpg1252\cocoartf1671\cocoasubrtf500
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Ejecutar el servidor NGINX:\
\
sudo nginx\
\
Ir a la ruta indicada:\
cd /User/cristhianimba/Sitio\
\
Ejecutar el ambiente virtual:\
source Projet1-env/bin/activate\
\
Cambiar de directorio para ejecutar el servidor Gunicorn:\
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0
\cf0 cd /User/cristhianimba/Sitio/Projet1\
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0
\cf0  \
Ejecutar el Servidor Gunicorn:\
sudo gunicorn -preload -c gunicorn.conf.py Projet1.wsgi \
\
\
En caso de necesitar detener el servidor NGINX se utiliza el comando:\
sudo nginx -s stop}