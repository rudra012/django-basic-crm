from django.conf.urls import include, url
from django.contrib import admin
from admin import index
from .admin import custom_admin
import SupremeApp
urlpatterns = [
    # Examples:
    # url(r'^$', 'Supreme.views.home', name='home'),
    # url(r'^blog/', include('blog.urls')),

    url(r'^admin/', include(admin.site.urls)),
    url(r'custom_admin/', include(custom_admin.urls)),
    # url(r'^$', include(admin.site.urls)),
    url(r'^$', index),
    url(r'SupremeApp/', include(SupremeApp.urls)),
]
