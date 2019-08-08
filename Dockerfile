FROM nginx:latest

COPY ./resources/nginx.conf /etc/nginx/conf.d/default.conf
COPY ./resources/certificates /etc/nginx/certificates

COPY ./dist/. /usr/share/nginx/html
COPY manifest.xml /usr/share/nginx/html/manifest.xml
COPY ./assets /usr/share/nginx/html/assets

EXPOSE 80 443