from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from django.contrib import admin
from django.conf import settings
from django.contrib.auth.views import logout


class CustomAdmin(admin.AdminSite):
    site_header = 'Supreme'
    site_title = 'Supreme'
    index_title = 'Supreme Module'


custom_admin = CustomAdmin(name='custom_admin')


@login_required
def index(request):
    site_header = site_title = 'Supreme'
    has_permission = True
    site_url = '/'
    # content = """<a href="/mnp/">MNP (Mobile Number Portability)</a>
    #             <a href="/churn_enq/">Churn Enquiry</a>
    #             """
    return render(request, 'index.html', context=locals())


@login_required
def custom_logout(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]

    if session_id in settings.FORM_SESSION:
        settings.FORM_SESSION.pop(session_id)

    # if session_id in SESSION_PAGES:
    #     SESSION_PAGES.pop(session_id)

    return logout(request, next_page='/')
