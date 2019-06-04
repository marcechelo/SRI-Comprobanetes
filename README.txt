Ejecutar el servidor NGINX:

sudo nginx

Ir a la ruta indicada:
cd /User/cristhianimba/Sitio

Ejecutar el ambiente virtual:
source Projet1-env/bin/activate

Cambiar de directorio para ejecutar el servidor Gunicorn:
cd /User/cristhianimba/Sitio/Projet1

Ejecutar el Servidor Gunicorn:
sudo gunicorn -preload -c gunicorn.conf.py Projet1.wsgi 


En caso de necesitar detener el servidor NGINX se utiliza el comando:
sudo nginx -s stop