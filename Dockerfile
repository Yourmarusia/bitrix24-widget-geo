FROM nginx:alpine

COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY widget/   /usr/share/nginx/html/widget/
COPY cities_russia.json /usr/share/nginx/html/

EXPOSE 80
