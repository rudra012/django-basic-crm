from django.contrib import admin
# Register your models here.
from SupremeApp.models import SupremeModel


class SupremeAdmin(admin.ModelAdmin):
    fieldsets = [
        ('',
         dict(fields=[
             ('caf_num', 'cust_name', 'mdn_no', 'rate_plan', 'otaf_date',),
             ('account_balance', 'no_of_payments_made', 'line_1', 'line_2', 'city', 'cluster'),
             ('alternate_landline_number', 'alternate_mobile_number', 'email_id', 'bill_cycle', 'bill_delivery_mode'),

         ], )),
        ('',
         dict(fields=[
             ('final_calling_code', 'final_followup_date', 'final_calling_remarks',),
         ], )),
    ]
    readonly_fields = (
    'caf_num', 'cust_name', 'mdn_no', 'rate_plan', 'otaf_date', 'account_balance', 'no_of_payments_made', 'line_1',
    'line_2', 'city', 'cluster', 'alternate_landline_number', 'alternate_mobile_number', 'email_id', 'bill_cycle',
    'bill_delivery_mode')
    list_display = readonly_fields


admin.site.register(SupremeModel, SupremeAdmin)
