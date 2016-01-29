from django.contrib import admin
# Register your models here.
from SupremeApp.models import SupremeModel
import datetime


class SupremeAdmin(admin.ModelAdmin):
    fieldsets = [
        ('',
         dict(fields=[
             ('caf_num', 'account_balance',),
             ('cust_name', 'no_of_payments_made',),
             ('mdn_no', 'rate_plan',),
             ('alternate_landline_number', 'otaf_date',),
             ('alternate_mobile_number', 'bill_cycle',),
             ('address', 'cluster',),
             ('email_id', 'bill_delivery_mode'),
             ('attempt',),
         ], )),
        ('',
         dict(fields=[
             ('final_calling_code', 'final_followup_date', 'final_calling_remarks',),
         ], )),
    ]
    readonly_fields = (
        'caf_num', 'cust_name', 'mdn_no', 'rate_plan', 'otaf_date', 'account_balance', 'no_of_payments_made', 'address',
        'cluster', 'alternate_landline_number', 'alternate_mobile_number', 'email_id', 'bill_cycle', 'attempt',
        'bill_delivery_mode')
    list_display = ('processed', 'cust_name', 'caf_num', 'mdn_no', 'final_tc_name', 'status',
                    'final_calling_date', 'final_calling_code', 'final_followup_date', 'final_calling_remarks',
                    'tc_1_attempt_date', 'tc_1_attempt_code', 'tc_1_attempt_remarks',
                    'tc_2_attempt_date', 'tc_2_attempt_code', 'tc_2_attempt_remarks',
                    'tc_3_attempt_date', 'tc_3_attempt_code', 'tc_3_attempt_remarks',
                    'tc_4_attempt_date', 'tc_4_attempt_code', 'tc_4_attempt_remarks',
                    'tc_5_attempt_date', 'tc_5_attempt_code', 'tc_5_attempt_remarks',
                    'tc_6_attempt_date', 'tc_6_attempt_code', 'tc_6_attempt_remarks',
                    )
    search_fields = ('mdn_no', 'cust_name',)
    date_hierarchy = 'date_created'
    list_display_links = ('cust_name',)

    def save_model(self, request, obj, form, change):
        print form.data
        obj.attempt = int(obj.attempt) + 1
        print obj.attempt
        attempt_date = 'tc_%s_attempt_date' % obj.attempt
        print attempt_date
        obj.__setattr__('tc_{}_name'.format(obj.attempt), str(request.user))
        obj.__setattr__('tc_{}_attempt_date'.format(obj.attempt), datetime.datetime.now())
        obj.__setattr__('tc_{}_attempt_code'.format(obj.attempt), form.data['final_calling_code'])
        obj.__setattr__('tc_{}_attempt_remarks'.format(obj.attempt), form.data['final_calling_remarks'])
        obj.processed = True
        obj.final_tc_name = str(request.user)
        obj.save()

    class Media:
        css = {'all': ('css/no-addanother-button.css',)}

    def has_add_permission(self, request):
        return False

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser

    def get_list_filter(self, request):
        return self.list_filter if request.user.is_superuser else ()

    def get_queryset(self, request):
        qs = super(SupremeAdmin, self).get_queryset(request)
        url_query = bool(request.GET.get('q'))
        url_query = url_query if '_changelist_filters' not in request.GET else True
        if not request.user.is_superuser and not url_query:
            qs = qs.none()
        return qs


admin.site.register(SupremeModel, SupremeAdmin)
