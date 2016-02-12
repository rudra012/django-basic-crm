from django import forms
import datetime
from Supreme.widgets import SelectDateWidget
from bootstrap3_datepicker.fields import DatePickerField
from bootstrap3_datepicker.widgets import DatePickerInput

from .models import SupremeModel


class UploadFileForm(forms.Form):
    file = forms.FileField()
    sheet_no = forms.IntegerField(min_value=1, initial=1)


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



