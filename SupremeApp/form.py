from django import forms
import datetime
from Supreme.widgets import SelectDateWidget


class UploadFileForm(forms.Form):
    file = forms.FileField()
    sheet_no = forms.IntegerField(min_value=1)


class DownloadFileForm(forms.Form):
    from_date = forms.DateField(required=True, widget=SelectDateWidget, initial=datetime.date.today())
    to_date = forms.DateField(required=True, widget=SelectDateWidget, initial=datetime.date.today())
    based_on = forms.ChoiceField(choices=[('Last Modified',) * 2,
                                          ('Create Time',) * 2])
