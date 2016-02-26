import datetime
from django.conf import settings
from django.contrib import admin
from django.db import models
from string import join

from .models import *
from django.db.models import Q
from django.forms import TextInput, Textarea
from datetimewidget.widgets import DateTimeWidget
from django.contrib import messages

from SupremeApp.models import SupremeModel, TCModel


class TCaddInline(admin.TabularInline):
    model = TCModel
    fields = ("calling_code", "followup_date", "calling_remarks",)

    def save_model(self, request, obj, form, change):
        # print "SAVE inline MODEL", dir(obj)
        # print form.data
        obj.calling_date = datetime.datetime.now()
        obj.tc_name = str(request.user)
        obj.save()

    def get_max_num(self, request, obj=None, **kwargs):
        max_num = 1
        return max_num

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False

    dateTimeOptions = {
        'format': 'yyyy-mm-dd hh:ii:00',
        'pickerPosition': 'top-right',
        'minuteStep': 5,
        'showMeridian': False,
        'todayHighlight': True,
        'autoclose': True
    }
    formfield_overrides = {
        models.DateTimeField: {'widget': DateTimeWidget(options=dateTimeOptions,)},
    }


class TCreadInline(admin.TabularInline):
    model = TCModel
    fields = ("calling_code", "followup_date", "calling_remarks", "tc_name", "calling_date")
    readonly_fields = fields

    def has_add_permission(self, request):
        return False

    def has_delete_permission(self, request, obj=None):
        return False


class SupremeAdmin(admin.ModelAdmin):
    # form = SupremeAdminForm
    inlines = [TCaddInline, TCreadInline]
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
        # ('',
        #  dict(fields=[
        #      ('calling_code', 'followup_date', 'calling_remarks',),
        #  ], )),
        ('TC summery',
         dict(fields=[
             ('final_calling_code', 'final_followup_date', 'final_calling_remarks', 'final_calling_date',),
         ],
         )),
    ]
    readonly_fields = (
        'caf_num', 'cust_name', 'mdn_no', 'rate_plan', 'otaf_date', 'account_balance', 'no_of_payments_made', 'address',
        'cluster', 'alternate_landline_number', 'alternate_mobile_number', 'email_id', 'bill_cycle', 'attempt',
        'bill_delivery_mode',
        'final_calling_date', 'final_calling_code', 'final_followup_date', 'final_calling_remarks',
    )
    list_display = (
        'processed', 'cust_name', 'caf_num', 'mdn_no', 'final_tc_name', 'status', 'pending_amt', 'account_balance',
        'final_calling_date', 'final_calling_code', 'final_followup_date', 'final_calling_remarks',
    )
    search_fields = ('mdn_no', 'cust_name',)
    date_hierarchy = 'date_created'
    list_display_links = ('cust_name',)
    ordering = ('processed', '-final_followup_date')
    list_filter = ('processed', 'bill_cycle', 'allocation_date')
    # save_on_top = True

    def save_formset(self, request, form, formset, change):
        instances = formset.save(commit=False)
        for instance in instances:
            # print dir(instance)
            instance.tc_name = str(request.user)
            instance.calling_date = datetime.datetime.now()
            instance.save()
        formset.save()

    def save_model(self, request, obj, form, change):
        # print "SAVE MODEL", dir(obj)
        # print dir(obj.tcmodel_set)
        # print form.data
        # attempt_date = 'tc_%s_attempt_date' % obj.attempt
        # print attempt_date
        # obj.__setattr__('tc_{}_name'.format(obj.attempt), str(request.user))
        # obj.__setattr__('tc_{}_attempt_date'.format(obj.attempt), datetime.datetime.now())
        # obj.__setattr__('tc_{}_attempt_code'.format(obj.attempt), form.data['calling_code'])
        # obj.__setattr__('final_calling_code'.format(obj.attempt), form.data['calling_code'])
        # obj.__setattr__('tc_{}_attempt_remarks'.format(obj.attempt), form.data['calling_remarks'])
        # obj.__setattr__('tc_{}_attempt_followup_0'.format(obj.attempt), form.data['followup_date_0'])
        # obj.__setattr__('tc_{}_attempt_followup_1'.format(obj.attempt), form.data['followup_date_1'])
        # obj.__setattr__('tc_{}_attempt_followup'.format(obj.attempt), obj.followup_date)
        # obj.__setattr__('tcmodel_set-0-followup_date', form.data['followup_date_1'])
        # obj.final_followup_date = obj.followup_date
        # obj.tcmodel_set-0-followup_date = obj.followup_date

        # obj.__setattr__('final_calling_date', datetime.datetime.now())
        # form.data.pop("followup_date_0")
        # form.data.pop("followup_date_1")
        # form.cleaned_data.pop("followup_date_0")
        # form.cleaned_data.pop("followup_date")
        # obj.followup_date = None
        # obj.calling_remarks = None
        # obj.calling_code = None
        # print type(form.data['tcmodel_set-0-followup_date_0'])
        # print type(form.data['tcmodel_set-0-followup_date_1'])
        final_followup_date = form.data['tcmodel_set-0-followup_date']
        obj.attempt = int(obj.attempt) + 1
        obj.__setattr__('final_tc_name', str(request.user))
        obj.__setattr__('final_calling_remarks', form.data['tcmodel_set-0-calling_remarks'])
        obj.__setattr__('final_calling_code', form.data['tcmodel_set-0-calling_code'])
        if final_followup_date:
            obj.final_followup_date = final_followup_date
        obj.final_calling_date = datetime.datetime.now()
        obj.final_tc_name = str(request.user)
        obj.processed = True
        obj.save()

    class Media:
        # css = {'all': (
        #     settings.BASE_DIR + '/static/css/no-addanother-button.css',
        #     settings.BASE_DIR + '/static/admin/css/forms.css'), }

        js = ['/static/admin/js/test.js']

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
        # print qs
        # processed_qs = filter(lambda x: x.processed, qs)
        # call_back_qs = filter(lambda x: x.final_followup_date, qs)
        # fmt = '%Y-%m-%d %H:%M:%S'
        make_above_pks = []
        for model_obj in qs:
            if model_obj.final_followup_date:
                time_dif = abs((model_obj.final_followup_date - datetime.datetime.now()).total_seconds() / 60)
                # print time_dif
                if time_dif < 5:
                    make_above_pks.append(model_obj.pk)
        # print make_above_pks, "MAKE ABOVE"
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


class SettingAdmin(admin.ModelAdmin):
    def has_add_permission(self, request):
        return not Setting.objects.exists()

    def has_delete_permission(self, request, obj=None):
        return False

    list_display = ('title', 'active_days', 'time', 'no_of_backup_days')
    fields = ['title', 'days', 'time', 'no_of_backup_days']
    if Setting.objects.exists():
        readonly_fields = ['title']

    def save_model(self, request, obj, form, change):
        messages.warning(request, 'Note:- These changes will apply from tomorrow onwords.')
        super(SettingAdmin, self).save_model(request, obj, form, change)

    def active_days(self, instance):
        lst = []
        if isinstance(instance.days, list):
            for v in instance.days:
                lst.append(dict(instance.days_list)[v])
            return join(lst, ", ")
        else:
            return "-"
    dateTimeOptions = {
        'format': 'hh:ii:00',
        'pickerPosition': 'top-right',
        'minuteStep': 5,
        'startDate': datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d'),
        'showMeridian': False,
        'startView': 1,
        'maxView': 1,
        'autoclose': True
    }

    formfield_overrides = {
        models.TimeField: {'widget': DateTimeWidget(options=dateTimeOptions,)},
    }


class UserDetailAdmin(admin.ModelAdmin):
    fields = ('name', 'email')
    list_display = ('name', 'email')


admin.site.register(SupremeModel, SupremeAdmin)
admin.site.register(UserDetail, UserDetailAdmin)
admin.site.register(Setting, SettingAdmin)
