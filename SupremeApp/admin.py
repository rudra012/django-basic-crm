from django.contrib import admin
# Register your models here.
from SupremeApp.models import SupremeModel
import datetime
from django.conf import settings
from django.db import models
from django.forms import TextInput, Textarea


class SupremeAdmin(admin.ModelAdmin):
    formfield_overrides = {
        models.CharField: {'widget': TextInput(attrs={'size': '25', 'width': '50%'})},
        models.TextField: {'widget': Textarea(attrs={'rows': 4, 'cols': 40})},
    }
    fieldsets = [
        ('Customer Details',
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
             ('calling_code', 'followup_date', 'calling_remarks',),
         ], )),
        ('TC summery',
         dict(fields=[
             ('tc_1_attempt_date', 'tc_1_attempt_code', 'tc_1_attempt_remarks',),
             ('tc_2_attempt_date', 'tc_2_attempt_code', 'tc_2_attempt_remarks',),
             ('tc_3_attempt_date', 'tc_3_attempt_code', 'tc_3_attempt_remarks',),
             ('tc_4_attempt_date', 'tc_4_attempt_code', 'tc_4_attempt_remarks',),
             ('tc_5_attempt_date', 'tc_5_attempt_code', 'tc_5_attempt_remarks',),
             ('tc_6_attempt_date', 'tc_6_attempt_code', 'tc_6_attempt_remarks',),
             ('final_calling_date', 'final_calling_code', 'final_calling_remarks', 'final_followup_date',),
         ],
         )),
    ]
    readonly_fields = (
        'caf_num', 'cust_name', 'mdn_no', 'rate_plan', 'otaf_date', 'account_balance', 'no_of_payments_made', 'address',
        'cluster', 'alternate_landline_number', 'alternate_mobile_number', 'email_id', 'bill_cycle', 'attempt',
        'bill_delivery_mode', 'tc_1_attempt_date', 'tc_1_attempt_code', 'tc_1_attempt_remarks', 'tc_2_attempt_date',
        'tc_2_attempt_code', 'tc_2_attempt_remarks', 'tc_3_attempt_date', 'tc_3_attempt_code', 'tc_3_attempt_remarks',
        'tc_4_attempt_date', 'tc_4_attempt_code', 'tc_4_attempt_remarks', 'tc_5_attempt_date', 'tc_5_attempt_code',
        'tc_5_attempt_remarks', 'tc_6_attempt_date', 'tc_6_attempt_code', 'tc_6_attempt_remarks',
        'final_calling_date', 'final_calling_code', 'final_followup_date', 'final_calling_remarks',
    )
    list_display = (
        'processed', 'cust_name', 'caf_num', 'mdn_no', 'final_tc_name', 'status',
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
    ordering = ('processed', '-final_followup_date')
    list_filter = ('processed', 'final_calling_code')

    def save_model(self, request, obj, form, change):
        print "SAVE MODEL"
        # print form.data
        obj.attempt = int(obj.attempt) + 1
        # print obj.attempt
        # attempt_date = 'tc_%s_attempt_date' % obj.attempt
        # print attempt_date
        obj.__setattr__('final_tc_name'.format(obj.attempt), str(request.user))
        obj.__setattr__('tc_{}_name'.format(obj.attempt), str(request.user))
        obj.__setattr__('tc_{}_attempt_date'.format(obj.attempt), datetime.datetime.now())
        obj.__setattr__('tc_{}_attempt_code'.format(obj.attempt), form.data['calling_code'])
        obj.__setattr__('final_calling_code'.format(obj.attempt), form.data['calling_code'])
        obj.__setattr__('tc_{}_attempt_remarks'.format(obj.attempt), form.data['calling_remarks'])
        obj.__setattr__('final_calling_remarks', form.data['calling_remarks'])
        # obj.__setattr__('tc_{}_attempt_followup_0'.format(obj.attempt), form.data['followup_date_0'])
        # obj.__setattr__('tc_{}_attempt_followup_1'.format(obj.attempt), form.data['followup_date_1'])
        obj.__setattr__('tc_{}_attempt_followup'.format(obj.attempt), obj.followup_date)
        # obj.__setattr__('final_followup_date_0', form.data['followup_date_0'])
        # obj.__setattr__('final_followup_date_1', form.data['followup_date_1'])
        obj.final_followup_date = obj.followup_date
        obj.__setattr__('final_calling_date', datetime.datetime.now())
        # form.data.pop("followup_date_0")
        # form.data.pop("followup_date_1")
        # form.cleaned_data.pop("followup_date_0")
        # form.cleaned_data.pop("followup_date")
        obj.followup_date = None
        obj.calling_remarks = None
        obj.calling_code = None
        obj.processed = True
        obj.final_tc_name = str(request.user)
        obj.save()

    class Media:
        css = {'all': ('css/no-addanother-button.css', settings.BASE_DIR + '/static/admin/css/forms.css')}

    # class Media:
    #     def __init__(self):
    #         pass
    #     css = {'all': ('css/no-addanother-button.css',
    #                    # settings.BASE_DIR + '/static/admin/css/forms.css'
    #                    )}
    #     from django.conf import settings
    #     static_url = getattr(settings, 'STATIC_URL', '/static/')
    #     js = ['http://ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js', static_url + 'admin/js/test.js']

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
            # For Agent
            qs = SupremeModel.objects.filter(final_tc_name=request.user).filter(~Q(status="Paid"))
        else:
            # For master
            return qs
        # print map(lambda x: x.processed, qs)
        # print map(lambda x: x.final_followup_date, qs)
        print qs
        # processed_qs = filter(lambda x: x.processed, qs)
        # call_back_qs = filter(lambda x: x.final_followup_date, qs)
        # fmt = '%Y-%m-%d %H:%M:%S'
        make_above_pks = []
        for model_obj in qs:
            if model_obj.final_followup_date:
                time_dif = abs((model_obj.final_followup_date - datetime.datetime.now()).total_seconds() / 60)
                print time_dif
                if time_dif < 5:
                    make_above_pks.append(model_obj.pk)
        print make_above_pks, "MAKE ABOVE"
        if make_above_pks:
            qs = qs.filter(pk__in=make_above_pks)
        # print dir(qs)
        # for query_set in make_above_qs:
        #     qs.delete(query_set)
        #     qs.append(qs)
        # for query_set in call_back_qs:
        #     time_dif = abs((query_set.final_followup_date - datetime.datetime.now()).total_seconds() / 60)
        #     print time_dif
        #     if time_dif < 5:
        #         make_above_qs.append(query_set)
        # none_processed_qs = filter(lambda x: not x.processed, qs)
        # print(processed_qs, none_processed_qs)
        # print call_back_qs
        # print sorted(qs, key=lambda x: x.processed)
        # print qs
        return qs


admin.site.register(SupremeModel, SupremeAdmin)
