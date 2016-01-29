from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from django.contrib import admin


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
