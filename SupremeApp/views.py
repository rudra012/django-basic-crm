import traceback

import datetime
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect
from django.shortcuts import render
import xlrd

from SupremeApp.form import UploadFileForm
from SupremeApp.models import SupremeModel


def xmldate_to_pdate(xdate):
    """
    Converts xml date type to python date
    :return:
    """
    try:
        return datetime.date(*xlrd.xldate_as_tuple(xdate, 0))
    except:
        return None


def handle_excel_file(ufile, sheet_no):
    print('reading file', ufile)
    wb = xlrd.open_workbook(file_contents=ufile.read())
    print wb.nsheets
    if sheet_no not in range(1, wb.nsheets + 1):
        return ["Sheet number {} not in range. There are {} sheets".format(sheet_no, wb.nsheets)]

    sheet = wb.sheet_by_index(sheet_no - 1)
    msg = []
    uploaded = []
    try:
        for i in range(1, sheet.nrows):
            print(i)

            for col in range(sheet.ncols):
                print(col, sheet.cell(i, col).value, type(sheet.cell(i, col)), sheet.name)

            try:
                SupremeModel(
                    caf_num='%.0f' % sheet.cell(i, 0).value,
                    cust_name=sheet.cell(i, 1).value,
                    customer_category=sheet.cell(i, 2).value,
                    risk_class_code=sheet.cell(i, 3).value,
                    customer_type=sheet.cell(i, 4).value,
                    customer_ecl=sheet.cell(i, 5).value,
                    mdn_no='%.0f' % sheet.cell(i, 6).value,
                    adc_status=sheet.cell(i, 7).value,
                    service_type=sheet.cell(i, 8).value,
                    rate_plan=sheet.cell(i, 9).value,
                    otaf_date=xmldate_to_pdate(sheet.cell(i, 10).value),
                    phongen_status=sheet.cell(i, 11).value,
                    desired_service=sheet.cell(i, 12).value,
                    og_bar=sheet.cell(i, 13).value,
                    account_balance=sheet.cell(i, 14).value,
                    pending_amount=sheet.cell(i, 15).value,
                    debtors_age=sheet.cell(i, 16).value,
                    voluntary_deposit=sheet.cell(i, 17).value,
                    unbilled_ild=sheet.cell(i, 18).value,
                    total_unbilled=sheet.cell(i, 19).value,
                    billed_not_due_amount=sheet.cell(i, 20).value,
                    b1=sheet.cell(i, 21).value,
                    b2=sheet.cell(i, 22).value,
                    b3=sheet.cell(i, 23).value,
                    b4=sheet.cell(i, 24).value,
                    b5=sheet.cell(i, 25).value,
                    b6=sheet.cell(i, 26).value,
                    b7=sheet.cell(i, 27).value,
                    b8=sheet.cell(i, 28).value,
                    billed_and_overdue_amount=sheet.cell(i, 29).value,
                    billed_outstanding_amount=sheet.cell(i, 30).value,
                    exposure=sheet.cell(i, 31).value,
                    latest_bill_due_date=xmldate_to_pdate(sheet.cell(i, 32).value),
                    no_of_invoice_raised=sheet.cell(i, 33).value,
                    total_invoice_amout=sheet.cell(i, 34).value,
                    no_of_payments_made=sheet.cell(i, 35).value,
                    payments_made_till_date=sheet.cell(i, 36).value,
                    total_adjustment_amount=sheet.cell(i, 37).value,
                    dapo_amount=sheet.cell(i, 38).value,
                    ciou_terr=sheet.cell(i, 39).value,
                    zone_terr=sheet.cell(i, 40).value,
                    cluster_terr=sheet.cell(i, 41).value,
                    circle_terr=sheet.cell(i, 42).value,
                    country_terr=sheet.cell(i, 43).value,
                    deposit=sheet.cell(i, 44).value,
                    no_of_active_services='%.0f' % sheet.cell(i, 45).value,
                    ecs=sheet.cell(i, 46).value,
                    daf_num=sheet.cell(i, 47).value,
                    ciou_name=sheet.cell(i, 48).value,
                    zone_name=sheet.cell(i, 49).value,
                    cluster_name=sheet.cell(i, 50).value,
                    circle_name=sheet.cell(i, 51).value,
                    line_1=sheet.cell(i, 52).value,
                    line_2=sheet.cell(i, 53).value,
                    city=sheet.cell(i, 54).value,
                    cluster=sheet.cell(i, 55).value,
                    state=sheet.cell(i, 56).value,
                    country=sheet.cell(i, 57).value,
                    postcode=sheet.cell(i, 58).value,
                    last_payment_date=xmldate_to_pdate(sheet.cell(i, 59).value),
                    excltype=sheet.cell(i, 60).value,
                    child=sheet.cell(i, 61).value,
                    do_not_disturb=sheet.cell(i, 62).value,
                    handset_model=sheet.cell(i, 63).value,
                    oldest_open_inv_due_date=xmldate_to_pdate(sheet.cell(i, 64).value),
                    last_bill_issue_date=xmldate_to_pdate(sheet.cell(i, 65).value),
                    last_payment_amount=sheet.cell(i, 67).value,
                    no_buf_debt_age='%.0f' % sheet.cell(i, 68).value,
                    data_unbilled_amount=sheet.cell(i, 69).value,
                    rconnect_plan=sheet.cell(i, 70).value,
                    bb_plan=sheet.cell(i, 71).value,
                    premier_category=sheet.cell(i, 72).value,
                    service_subtype=sheet.cell(i, 73).value,
                    last_month_score=sheet.cell(i, 74).value,
                    last2_month_score=sheet.cell(i, 75).value,
                    last3_month_score=sheet.cell(i, 76).value,
                    last4_month_score=sheet.cell(i, 77).value,
                    last5_month_score=sheet.cell(i, 78).value,
                    last6_month_score=sheet.cell(i, 79).value,
                    curr_payment_behaviour_class=sheet.cell(i, 80).value,
                    prev_payment_behaviour_class=sheet.cell(i, 81).value,
                    alternate_landline_number='%.0f' % sheet.cell(i, 82).value,
                    alternate_mobile_number='%.0f' % sheet.cell(i, 83).value,
                    email_id=sheet.cell(i, 84).value,
                    corp_cust_category=sheet.cell(i, 85).value,
                    mnp_flag=sheet.cell(i, 86).value,
                    bill_cycle=sheet.cell(i, 87).value,
                    bill_delivery_mode=sheet.cell(i, 88).value,
                    company_name=sheet.cell(i, 89).value,
                    barring_reason=sheet.cell(i, 90).value,
                    barring_date=xmldate_to_pdate(sheet.cell(i, 91).value),
                    allocation_date=xmldate_to_pdate(sheet.cell(i, 92).value),
                    closing_date=xmldate_to_pdate(sheet.cell(i, 93).value),
                ).save()
                uploaded.append(str(sheet.cell(i, 6).value))
            except Exception, e:
                print traceback.format_exc()
    except Exception, e:
        print traceback.format_exc()
        return ["System Error: " + str(e)]

    if uploaded:
        msg.append("Following entries Uploaded: {}".format(', '.join(uploaded)))
    else:
        msg.append("No Entries uploaded")
    return msg


@login_required
def upload(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "UPLOAD REQ"
    messages = []
    print 'rm', request.method
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            ufile = request.FILES['file']
            if ufile.name.endswith('.xls') or ufile.name.endswith('.xlsx'):
                msg = handle_excel_file(ufile, int(form.data['sheet_no']))
            else:
                msg = [" .xls, .xlsx and .csv file formats are only supported."]
            messages = msg
        else:
            messages.append("Supply appropriate data.")
    form = UploadFileForm()
    return render(request, 'SupremeApp/upload.html', locals())


@login_required
def download(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "Download REQ"
    return render(request, 'SupremeApp/download.html', locals())


@login_required
def index(request):
    print "INDEX REQ", request
    return render(request, 'SupremeApp/index.html', locals())


@login_required
def data(request):
    print 'data req', request
    return HttpResponseRedirect("/admin/SupremeApp/suprememodel/")
