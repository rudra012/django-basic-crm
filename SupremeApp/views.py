import datetime
import os
import traceback
import xlrd
import xlsxwriter
from django.conf import settings
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.http import HttpResponseRedirect
from django.shortcuts import render
from SupremeApp.form import UploadFileForm, DownloadFileForm, RDownloadFileForm
from SupremeApp.models import SupremeModel, TCModel


def try_to_int(idata):
    try:
        return int(float(idata))
    except:
        return 0


def try_to_str_int(sidata):
    try:
        return '%.0f' % sidata
    except:
        return sidata


def xmldate_to_pdate(xdate):
    """
    Converts xml date type to python date
    :return:
    """
    try:
        return datetime.date(*xlrd.xldate_as_tuple(xdate, 0))
    except:
        return None


def handle_data_excel_file(ufile, sheet_no):
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

            # for col in range(sheet.ncols):
            #     print(col, sheet.cell(i, col).value, type(sheet.cell(i, col)), sheet.name)

            try:
                mobile_no = try_to_str_int(sheet.cell(i, 6).value)
                SupremeModel(
                    caf_num=try_to_str_int(sheet.cell(i, 0).value),
                    cust_name=sheet.cell(i, 1).value,
                    customer_category=sheet.cell(i, 2).value,
                    risk_class_code=sheet.cell(i, 3).value,
                    customer_type=sheet.cell(i, 4).value,
                    customer_ecl=sheet.cell(i, 5).value,
                    mdn_no=('%.0f' % sheet.cell(i, 6).value)[-10:],
                    adc_status=sheet.cell(i, 7).value,
                    service_type=sheet.cell(i, 8).value,
                    rate_plan=sheet.cell(i, 9).value,
                    otaf_date=xmldate_to_pdate(sheet.cell(i, 10).value),
                    phongen_status=sheet.cell(i, 11).value,
                    desired_service=sheet.cell(i, 12).value,
                    og_bar=sheet.cell(i, 13).value,
                    account_balance=sheet.cell(i, 14).value,
                    debtors_age=sheet.cell(i, 15).value,
                    voluntary_deposit=sheet.cell(i, 16).value,
                    unbilled_ild=sheet.cell(i, 17).value,
                    total_unbilled=sheet.cell(i, 18).value,
                    billed_not_due_amount=sheet.cell(i, 19).value,
                    b1=sheet.cell(i, 20).value,
                    b2=sheet.cell(i, 21).value,
                    b3=sheet.cell(i, 22).value,
                    b4=sheet.cell(i, 23).value,
                    b5=sheet.cell(i, 24).value,
                    b6=sheet.cell(i, 25).value,
                    b7=sheet.cell(i, 26).value,
                    b8=sheet.cell(i, 27).value,
                    billed_and_overdue_amount=sheet.cell(i, 28).value,
                    billed_outstanding_amount=sheet.cell(i, 29).value,
                    exposure=sheet.cell(i, 30).value,
                    latest_bill_due_date=xmldate_to_pdate(sheet.cell(i, 31).value),
                    no_of_invoice_raised=sheet.cell(i, 32).value,
                    total_invoice_amout=sheet.cell(i, 33).value,
                    no_of_payments_made=sheet.cell(i, 34).value,
                    payments_made_till_date=sheet.cell(i, 35).value,
                    total_adjustment_amount=sheet.cell(i, 36).value,
                    dapo_amount=sheet.cell(i, 37).value,
                    ciou_terr=sheet.cell(i, 38).value,
                    zone_terr=sheet.cell(i, 39).value,
                    cluster_terr=sheet.cell(i, 40).value,
                    circle_terr=sheet.cell(i, 41).value,
                    country_terr=sheet.cell(i, 42).value,
                    deposit=sheet.cell(i, 43).value,
                    no_of_active_services=try_to_str_int(sheet.cell(i, 44).value),
                    ecs=sheet.cell(i, 45).value,
                    daf_num=sheet.cell(i, 46).value,
                    ciou_name=sheet.cell(i, 47).value,
                    zone_name=sheet.cell(i, 48).value,
                    cluster_name=sheet.cell(i, 49).value,
                    circle_name=sheet.cell(i, 50).value,
                    line_1=sheet.cell(i, 51).value,
                    line_2=sheet.cell(i, 52).value,
                    city=sheet.cell(i, 53).value,
                    cluster=sheet.cell(i, 54).value,
                    state=sheet.cell(i, 55).value,
                    country=sheet.cell(i, 56).value,
                    postcode=sheet.cell(i, 57).value,
                    last_payment_date=xmldate_to_pdate(sheet.cell(i, 58).value),
                    excltype=sheet.cell(i, 59).value,
                    child=sheet.cell(i, 60).value,
                    do_not_disturb=sheet.cell(i, 61).value,
                    handset_model=sheet.cell(i, 62).value,
                    oldest_open_inv_due_date=xmldate_to_pdate(sheet.cell(i, 63).value),
                    last_bill_issue_date=xmldate_to_pdate(sheet.cell(i, 64).value),
                    last_biil_amount=xmldate_to_pdate(sheet.cell(i, 65).value),
                    last_payment_amount=sheet.cell(i, 66).value,
                    no_buf_debt_age=try_to_str_int(sheet.cell(i, 67).value),
                    data_unbilled_amount=sheet.cell(i, 68).value,
                    rconnect_plan=sheet.cell(i, 69).value,
                    bb_plan=sheet.cell(i, 70).value,
                    premier_category=sheet.cell(i, 71).value,
                    service_subtype=sheet.cell(i, 72).value,
                    last_month_score=sheet.cell(i, 73).value,
                    last2_month_score=sheet.cell(i, 74).value,
                    last3_month_score=sheet.cell(i, 75).value,
                    last4_month_score=sheet.cell(i, 76).value,
                    last5_month_score=sheet.cell(i, 77).value,
                    last6_month_score=sheet.cell(i, 78).value,
                    curr_payment_behaviour_class=sheet.cell(i, 79).value,
                    prev_payment_behaviour_class=sheet.cell(i, 80).value,
                    alternate_landline_number=try_to_str_int(sheet.cell(i, 81).value),
                    alternate_mobile_number=try_to_str_int(sheet.cell(i, 82).value),
                    email_id=sheet.cell(i, 83).value,
                    corp_cust_category=sheet.cell(i, 84).value,
                    mnp_flag=sheet.cell(i, 85).value,
                    bill_cycle=sheet.cell(i, 86).value,
                    bill_delivery_mode=sheet.cell(i, 87).value,
                    company_name=sheet.cell(i, 88).value,
                    barring_reason=sheet.cell(i, 89).value,
                    barring_date=xmldate_to_pdate(sheet.cell(i, 90).value),
                    allocation_date=xmldate_to_pdate(sheet.cell(i, 91).value),
                    closing_date=xmldate_to_pdate(sheet.cell(i, 92).value),
                    final_tc_name=sheet.cell(i, 93).value,
                    address=", ".join([sheet.cell(i, 51).value, sheet.cell(i, 52).value, sheet.cell(i, 53).value]),
                ).save()
                uploaded.append(str(mobile_no))
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
                msg = handle_data_excel_file(ufile, int(form.data['sheet_no']))
            else:
                msg = [" .xls, .xlsx and .csv file formats are only supported."]
            messages = msg
        else:
            messages.append("Supply appropriate data.")
    form = UploadFileForm()
    return render(request, 'SupremeApp/upload.html', locals())


def handle_paid_excel_file(ufile, sheet_no):
    print('reading paid file', ufile)
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

            # for col in range(sheet.ncols):
            #     print(col, sheet.cell(i, col).value, type(sheet.cell(i, col).value), sheet.name)

            try:
                mobile_no = '%.0f' % sheet.cell(i, 1).value
                payment_amt = '%.0f' % sheet.cell(i, 4).value
                credit_limit = '%.0f' % sheet.cell(i, 9).value
                print mobile_no, payment_amt, credit_limit
                print type(mobile_no), type(payment_amt)
                paid_user_obj_list = list(SupremeModel.objects.filter(mdn_no=mobile_no))
                print paid_user_obj_list
                # Sort to get latest object first
                paid_user_obj_list.sort(key=lambda x: x.date_created, reverse=True)
                if paid_user_obj_list:
                    last_user_obj = paid_user_obj_list[0]
                    print "&&&&&&&&&&&&&&&&&&&"
                    remain_payment = try_to_int(last_user_obj.account_balance) - try_to_int(payment_amt)
                    print last_user_obj.account_balance, payment_amt, remain_payment
                    if remain_payment < 100:
                        last_user_obj.status = "Paid"
                    else:
                        last_user_obj.status = "Partial Paid"
                    last_user_obj.save()
                    uploaded.append(mobile_no)
            except Exception, e:
                print traceback.format_exc()
    except Exception, e:
        print traceback.format_exc()
        return ["System Error: " + str(e)]
    # msg.append("All Entries Updated")
    if uploaded:
        msg.append("Following entries Updated: {}".format(', '.join(uploaded)))
    else:
        msg.append("No Entries Updated")
    return msg


@login_required
def paid_upload(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "UPLOAD REQ"
    messages = []
    print 'rm', request.method
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            ufile = request.FILES['file']
            if ufile.name.endswith('.xls') or ufile.name.endswith('.xlsx'):
                msg = handle_paid_excel_file(ufile, int(form.data['sheet_no']))
            else:
                msg = [" .xls, .xlsx and .csv file formats are only supported."]
            messages = msg
        else:
            messages.append("Supply appropriate data.")
    form = UploadFileForm()
    return render(request, 'SupremeApp/paid_upload.html', locals())


def get_supreme_app_data(from_date, to_date, based_on):
    """
    This is to collect Supreme App data based on date given by admin
    :return:
    """
    print based_on
    if based_on == "Last Modified":
        supreme_data_queryset = SupremeModel.objects.filter(date_modified__range=(from_date, to_date))
    elif based_on == "Create Time":
        supreme_data_queryset = SupremeModel.objects.filter(date_created__range=(from_date, to_date))
    else:
        supreme_data_queryset = SupremeModel.objects.none()
    return supreme_data_queryset


class IncrementVar(object):
    def __init__(self, value=0):
        self.inc_var = value

    def get_inc_var(self):
        # print("getting value")
        self._inc_var += 1
        return self._inc_var

    def set_inc_var(self, value):
        # print("setting value %s" % value)
        self._inc_var = value

    inc_var = property(get_inc_var, set_inc_var)


def create_temp_xlsx_file(supreme_app_data):
    """
    This will create temp xlsx file and return that path
    :return:
    """
    file_path = '/tmp/temp1.xlsx'
    book = xlsxwriter.Workbook(file_path)
    sheet = book.add_worksheet('CHURN_ENQ REVERT')
    # sheet.set_tab_color('red')

    # Different formats for attractive xls file
    bold = book.add_format({'bold': True})
    red_font = book.add_format({'bold': True, 'font_color': 'red'})
    red_bg = book.add_format({'bg_color': 'red', 'border': 1})
    yellow_bg = book.add_format({'bg_color': 'yellow', 'border': 1, 'bold': True})
    orange_bg = book.add_format({'bg_color': 'orange', 'border': 1})
    green_bg = book.add_format({'bg_color': 'green', 'border': 1})

    date_format = book.add_format({'num_format': 'mm/dd/yyyy'})
    # time_format = book.add_format({'num_format': 'hh:mm:ss'})
    date_time_format = book.add_format({'num_format': 'mm/dd/yy hh:mm AM/PM'})
    date_time_format_green_bg = book.add_format({'num_format': 'mm/dd/yy hh:mm AM/PM', 'bg_color': 'green', 'border': 1})
    # date_time_format = book.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
    # date_time_format = book.add_format({'num_format': 'mm/dd/yy hh:mm:ss'})

    heading = [
        "CAF_NUM",
        "CUST_NAME",
        "MDN_NO",
        "RATE_PLAN",
        "OTAF_DATE",
        "ACCOUNT_BALANCE",
        "NO_OF_PAYMENTS_MADE",
        "line_1",
        "line_2",
        "city",
        "Cluster",
        "ALTERNATE_LANDLINE_NUMBER",
        "ALTERNATE_MOBILE_NUMBER",
        "EMAILID",
        "BILL_CYCLE",
        "BILL_DELIVERY_MODE",
        "TC_NAME",
        "Final Calling Date",
        "Final Calling Code",
        "Final Calling Remarks",
        "Final_Followup_Date",
        "Status",
    ]
    for i, h in enumerate(heading):
        sheet.write(0, i, h)

    for i, data in enumerate(supreme_app_data):
        j = i + 1
        i_var = IncrementVar(-1)
        print j, i_var
        sheet.write(j, i_var.inc_var, data.caf_num)
        sheet.write(j, i_var.inc_var, data.cust_name)
        sheet.write(j, i_var.inc_var, data.mdn_no)
        sheet.write(j, i_var.inc_var, data.rate_plan)
        sheet.write(j, i_var.inc_var, data.otaf_date)
        sheet.write(j, i_var.inc_var, data.account_balance)
        sheet.write(j, i_var.inc_var, data.no_of_payments_made)
        sheet.write(j, i_var.inc_var, data.line_1)
        sheet.write(j, i_var.inc_var, data.line_2)
        sheet.write(j, i_var.inc_var, data.city)
        sheet.write(j, i_var.inc_var, data.cluster)
        sheet.write(j, i_var.inc_var, data.alternate_landline_number)
        sheet.write(j, i_var.inc_var, data.alternate_mobile_number)
        sheet.write(j, i_var.inc_var, data.email_id)
        sheet.write(j, i_var.inc_var, data.bill_cycle)
        sheet.write(j, i_var.inc_var, data.bill_delivery_mode)
        sheet.write(j, i_var.inc_var, data.final_tc_name)
        sheet.write(j, i_var.inc_var, data.final_calling_date, date_time_format)
        sheet.write(j, i_var.inc_var, data.final_calling_code)
        sheet.write(j, i_var.inc_var, data.final_calling_remarks)
        sheet.write(j, i_var.inc_var, data.final_followup_date, date_time_format)
        sheet.write(j, i_var.inc_var, data.status)
        tc_details = TCModel.objects.filter(superme_key_id=data.id)
        print tc_details
        heading_end = len(heading)
        for tc_index, detail in enumerate(tc_details):
            print tc_index, detail
            i_var_2 = IncrementVar(heading_end - 1)
            sheet.write(0, i_var_2.inc_var + tc_index, "TC %sth Dispo Code" % (tc_index + 1), yellow_bg)
            sheet.write(0, i_var_2.inc_var + tc_index, "TC %sth CB/PTP Date" % (tc_index + 1), green_bg)
            sheet.write(0, i_var_2.inc_var + tc_index, "TC %sth Remarks" % (tc_index + 1), orange_bg)
            sheet.write(0, i_var_2.inc_var + tc_index, "TC %sth Calling Date" % (tc_index + 1), green_bg)
            heading_end += 3
            sheet.write(j, i_var.inc_var, detail.calling_code, yellow_bg)
            sheet.write(j, i_var.inc_var, detail.followup_date, date_time_format_green_bg)
            sheet.write(j, i_var.inc_var, detail.calling_remarks, orange_bg)
            sheet.write(j, i_var.inc_var, detail.calling_date, date_time_format_green_bg)

    book.close()
    return file_path


@login_required
def download(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "Download REQ"
    messages = []
    if request.method == 'POST':
        print "POST REQ"
        form = DownloadFileForm(request.POST)
        if not form.is_valid():
            settings.FORM_SESSION[session_id] = form
            return render(request, 'churn_enq/search.html', locals())
        else:
            from_date = form.cleaned_data['from_date']
            to_date = form.cleaned_data['to_date']
            based_on = form.data['based_on']
            print from_date, to_date, based_on
            supreme_app_data = get_supreme_app_data(datetime.datetime.combine(from_date, datetime.time(0, 0)),
                                                    datetime.datetime.combine(to_date, datetime.time(23, 59)),
                                                    based_on)
            if not supreme_app_data:
                messages.append(" No Search Results")
            else:
                path = create_temp_xlsx_file(supreme_app_data)
                response = HttpResponse(file(path, 'r').read())
                response['Content-Disposition'] = 'attachment;filename=SUPREME_DATA_from_{}_to_{}.xlsx'.format(
                    from_date, to_date)
                response['Content-Length'] = os.path.getsize(path)
                return response

    else:
        print "GET REQ"
        if session_id in settings.FORM_SESSION.keys():
            form = settings.FORM_SESSION[session_id]
        else:
            form = DownloadFileForm()
            # messages.append('Error, previous data not found. Please search again.')
    return render(request, 'SupremeApp/download.html', locals())


@login_required
def report_download(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "Required Download REQ"
    messages = []
    if request.method == 'POST':
        print "POST REQ"
        form = RDownloadFileForm(request.POST)
        if not form.is_valid():
            settings.FORM_SESSION[session_id] = form
            return render(request, 'churn_enq/search.html', locals())
        else:
            from_date = form.cleaned_data['from_date']
            to_date = form.cleaned_data['to_date']
            based_on = form.data['based_on']
            print from_date, to_date
            supreme_app_data = get_supreme_app_data(datetime.datetime.combine(from_date, datetime.time(0, 0)),
                                                    datetime.datetime.combine(to_date, datetime.time(23, 59)),
                                                    based_on)
            print supreme_app_data
            if not supreme_app_data:
                messages.append(" No Search Results")
            else:
                path = create_temp_xlsx_file(supreme_app_data)
                response = HttpResponse(file(path, 'r').read())
                response['Content-Disposition'] = 'attachment;filename=SUPREME_DATA_from_{}_to_{}.xlsx'.format(
                    from_date, to_date)
                response['Content-Length'] = os.path.getsize(path)
                return response

    else:
        print "GET REQ"
        if session_id in settings.FORM_SESSION.keys():
            form = settings.FORM_SESSION[session_id]
        else:
            form = RDownloadFileForm()
            # messages.append('Error, previous data not found. Please search again.')
    return render(request, 'SupremeApp/report_download.html', locals())


@login_required
def index(request):
    print "INDEX REQ", request
    return render(request, 'SupremeApp/index.html', locals())


@login_required
def data(request):
    print 'data req', request
    return HttpResponseRedirect("/admin/SupremeApp/suprememodel/")
