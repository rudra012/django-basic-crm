from django import forms
import datetime
from Supreme.widgets import SelectDateWidget
from bootstrap3_datepicker.fields import DatePickerField
from bootstrap3_datepicker.widgets import DatePickerInput

from .models import SupremeModel

import floppyforms as forms


class Slider(forms.RangeInput):
    min = 5
    max = 1000
    step = 5


class UploadFileForm(forms.Form):
    file = forms.FileField()
    sheet_no = forms.IntegerField(min_value=1, initial=1)
    based_on = forms.ChoiceField(choices=[('Fast Upload',) * 2,
                                          ('Normal Upload',) * 2])
    generate_new_report_list = ((True, 'Yes'), (False, 'No'))
    generate_report = forms.ChoiceField(choices=generate_new_report_list)
    speed = forms.IntegerField(min_value=10, max_value=1000, initial=100, widget=Slider)


class DownloadFileForm(forms.Form):
    from_date = forms.DateField(required=True, widget=DatePickerInput(format="%Y-%m-%d"), initial=datetime.date.today())
    to_date = forms.DateField(required=True,widget=DatePickerInput(format="%Y-%m-%d"), initial=datetime.date.today())
    based_on = forms.ChoiceField(choices=[('Last Modified',) * 2,
                                          ('Create Time',) * 2])


class RDownloadFileForm(forms.Form):
    from_date = forms.DateField(required=True, widget=DatePickerInput(format="%Y-%m-%d"), initial=datetime.date.today())
    to_date = forms.DateField(required=True, widget=DatePickerInput(format="%Y-%m-%d"), initial=datetime.date.today())
    # based_on = forms.ChoiceField(choices=[('Last Modified',) * 2,
    #                                       ('Create Time',) * 2])
