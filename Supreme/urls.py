from django.conf.urls import include, url
from django.contrib import admin
from admin import index, custom_admin, custom_logout
import SupremeApp

urlpatterns = [
    # Examples:
    # url(r'^$', 'Supreme.views.home', name='home'),
    # url(r'^blog/', include('blog.urls')),
    url(r'^accounts/login/$', 'django.contrib.auth.views.login'),
    url(r'^accounts/logout/$', custom_logout),
    url(r'^admin/logout/$', custom_logout),
    url(r'^admin/', include(admin.site.urls)),
    url(r'custom_admin/', include(custom_admin.urls)),
    # url(r'^$', include(admin.site.urls)),
    url(r'^$', index, name='index'),
    url(r'SupremeApp/', include(SupremeApp.urls)),
]
