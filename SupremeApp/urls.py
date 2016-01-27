from django.conf.urls import patterns, url

from views import download, upload, index, data

urlpatterns = patterns('',
                       url(r'^$', index),
                       url(r'^upload/$', upload),
                       url(r'^download/$', download),
                       url(r'^data/$', data),
                       )
