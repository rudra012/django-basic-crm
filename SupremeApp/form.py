from django import forms
import datetime
from Supreme.widgets import SelectDateWidget


class UploadFileForm(forms.Form):
    file = forms.FileField()
    sheet_no = forms.IntegerField(min_value=1, initial=1)
    based_on = forms.ChoiceField(choices=[('Fast Upload',) * 2,
                                          ('Normal Upload',) * 2])
    speed = forms.IntegerField(min_value=1, max_value=1000, initial=100)


class DownloadFileForm(forms.Form):
    from_date = forms.DateField(required=True, widget=SelectDateWidget, initial=datetime.date.today())
    to_date = forms.DateField(required=True, widget=SelectDateWidget, initial=datetime.date.today())
    based_on = forms.ChoiceField(choices=[('Last Modified',) * 2,
                                          ('Create Time',) * 2])


class RDownloadFileForm(forms.Form):
    from_date = forms.DateField(required=True, widget=SelectDateWidget, initial=datetime.date.today())
    to_date = forms.DateField(required=True, widget=SelectDateWidget, initial=datetime.date.today())
    # based_on = forms.ChoiceField(choices=[('Last Modified',) * 2,
    #                                       ('Create Time',) * 2])
