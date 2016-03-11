import os
import time
import traceback
from decimal import Decimal
import xlrd
import xlsxwriter
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.db.models import Q
from django.db.models import Sum
from django.http import HttpResponse
from django.http import HttpResponseRedirect
from django.shortcuts import render
from SupremeApp.form import *
from SupremeApp.models import *


def try_to_int(idata):
    try:
        return int(float(idata))
    except:
        return 0


def try_to_str_int(sidata):
    try:
        return '%.0f' % float(sidata)
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


def chunks(iterable, n):
    """Yield successive n-sized chunks from l."""
    for i in xrange(0, len(iterable), n):
        yield iterable[i:i + n]


def upload_data_excel_file(ufile, sheet_no, based_on, speed, user, generate_report):
    print('reading file', ufile)
    wb = xlrd.open_workbook(file_contents=ufile.read())
    print wb.nsheets
    if sheet_no not in range(1, wb.nsheets + 1):
        return [{"error": "Sheet number {} not in range. There are {} sheets".format(sheet_no, wb.nsheets)}]

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
                    mdn_no=mobile_no,
                    adc_status=sheet.cell(i, 7).value,
                    service_type=sheet.cell(i, 8).value,
                    rate_plan=sheet.cell(i, 9).value,
                    otaf_date=xmldate_to_pdate(sheet.cell(i, 10).value),
                    phongen_status=sheet.cell(i, 11).value,
                    desired_service=sheet.cell(i, 12).value,
                    og_bar=sheet.cell(i, 13).value,
                    account_balance=sheet.cell(i, 14).value,
                    pending_amt=sheet.cell(i, 14).value,
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
                    no_of_invoice_raised=try_to_str_int(sheet.cell(i, 32).value),
                    total_invoice_amout=sheet.cell(i, 33).value,
                    no_of_payments_made=try_to_str_int(sheet.cell(i, 34).value),
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
                    bill_cycle=try_to_str_int(sheet.cell(i, 86).value),
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
        return [{"error": "System Error: " + str(e)}]

    if uploaded:
        msg.append({"success": "Following entries Uploaded: {}".format(', '.join(uploaded))})
        u = UploadFileHistory()
        u.file_name = ufile
        u.sheet_no = sheet_no
        u.based_on = based_on
        u.speed = speed
        u.upload_type = "D"
        u.uploaded_date = datetime.datetime.now()
        u.no_of_records = len(uploaded)
        u.user = user
        if str(generate_report) == "True":
            u.generate_new_report = True
        else:
            u.generate_new_report = False
        u.save()
    else:
        msg.append({"warning": "No Entries uploaded"})
    return msg


def fast_upload_data_excel_file(ufile, sheet_no, based_on, speed, user, generate_report):
    print('speed reading file', ufile, speed)
    wb = xlrd.open_workbook(file_contents=ufile.read())
    print wb.nsheets
    if sheet_no not in range(1, wb.nsheets + 1):
        return [{"error": "Sheet number {} not in range. There are {} sheets".format(sheet_no, wb.nsheets)}]

    sheet = wb.sheet_by_index(sheet_no - 1)
    msg = []
    uploaded = []
    error_mobile_no = []
    error_mobile_no_range = []
    try:
        for range_list in chunks(range(1, sheet.nrows), speed):
            # print(range_list)
            # for col in range(sheet.ncols):
            #     print(col, sheet.cell(i, col).value, type(sheet.cell(i, col)), sheet.name)
            model_list = []
            try:
                for i in range_list:
                    mobile_no = ''
                    try:
                        # print sheet.cell(i, 6).value
                        mobile_no = try_to_str_int(sheet.cell(i, 6).value)
                        # print mobile_no, type(mobile_no)
                        model_obj = SupremeModel(
                            caf_num=try_to_str_int(sheet.cell(i, 0).value),
                            cust_name=sheet.cell(i, 1).value,
                            customer_category=sheet.cell(i, 2).value,
                            risk_class_code=sheet.cell(i, 3).value,
                            customer_type=sheet.cell(i, 4).value,
                            customer_ecl=sheet.cell(i, 5).value,
                            mdn_no=mobile_no[:10],
                            adc_status=sheet.cell(i, 7).value,
                            service_type=sheet.cell(i, 8).value,
                            rate_plan=sheet.cell(i, 9).value,
                            otaf_date=xmldate_to_pdate(sheet.cell(i, 10).value),
                            phongen_status=sheet.cell(i, 11).value,
                            desired_service=sheet.cell(i, 12).value,
                            og_bar=sheet.cell(i, 13).value,
                            account_balance=sheet.cell(i, 14).value,
                            pending_amt=sheet.cell(i, 14).value,
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
                            address=", ".join(
                                [str(sheet.cell(i, 51).value),
                                 str(sheet.cell(i, 52).value),
                                 str(sheet.cell(i, 53).value)]),
                        )
                        uploaded.append(str(mobile_no))
                        model_list.append(model_obj)
                    except Exception, e:
                        print traceback.format_exc()
                        error_mobile_no.append(str(mobile_no))
            except Exception, e:
                print traceback.format_exc()
                error_mobile_no_range.append(
                    str(sheet.cell(range_list[0], 6).value) + " : " + str(sheet.cell(range_list[-1], 6).value))
            # print model_list
            SupremeModel.objects.bulk_create(model_list)
    except Exception, e:
        print traceback.format_exc()
        return [{"error": "System Error: " + str(e)}]

    if uploaded:
        msg.append({"success": "Following entries Uploaded: {}".format(', '.join(uploaded))})
        u = UploadFileHistory()
        u.file_name = ufile
        u.sheet_no = sheet_no
        u.based_on = based_on
        u.speed = speed
        u.upload_type = "D"
        u.uploaded_date = datetime.datetime.now()
        u.no_of_records = len(uploaded)
        u.user = user
        if str(generate_report) == "True":
            u.generate_new_report = True
        else:
            u.generate_new_report = False
        u.save()
    else:
        msg.append({"warning": "No Entries uploaded"})

    if error_mobile_no:
        msg.append({"error": "Some problem in mobile numbers: {}".format(','.join(error_mobile_no))})

    if error_mobile_no_range:
        msg.append({"error": "Some problem in mobile range: {}".format(','.join(error_mobile_no_range))})

    return msg


@login_required
def upload(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "UPLOAD REQ"
    message_list = []
    print 'rm', request.method
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            ufile = request.FILES['file']
            sheel_no = int(form.data['sheet_no'])
            based_on = form.data['based_on']
            speed = int(form.data['speed'])
            generate_report = form.data['generate_report']
            if ufile.name.endswith('.xls') or ufile.name.endswith('.xlsx'):
                if based_on == "Normal Upload":
                    msg = upload_data_excel_file(ufile, sheel_no, based_on, speed, request.user, generate_report)
                else:
                    msg = fast_upload_data_excel_file(ufile, sheel_no, based_on, speed, request.user, generate_report)
            else:
                msg = [{"error": " .xls, .xlsx and .csv file formats are only supported."}]
            message_list = msg
        else:
            message_list.append({"error": "Supply appropriate data."})
    for msg in message_list:
        for k, v in msg.iteritems():
            if k == "warning":
                messages.warning(request, v)
            elif k == "error":
                print v
                messages.error(request, v)
            elif k == "success":
                messages.success(request, v)

    form = UploadFileForm()
    results = UploadFileHistory.objects.filter(upload_type='D').order_by('-uploaded_date').values()
    return render(request, 'SupremeApp/upload.html', locals())


def upload_paid_excel_file(ufile, sheet_no, based_on, speed, user):
    print('reading paid file', ufile)
    wb = xlrd.open_workbook(file_contents=ufile.read())
    # print wb.nsheets
    if sheet_no not in range(1, wb.nsheets + 1):
        return [{"error": "Sheet number {} not in range. There are {} sheets".format(sheet_no, wb.nsheets)}]

    sheet = wb.sheet_by_index(sheet_no - 1)
    msg = []
    uploaded = []
    not_found_caf = []
    error_caf_no = []
    try:
        for i in range(1, sheet.nrows):
            # print(i)

            # for col in range(sheet.ncols):
            #     print(col, sheet.cell(i, col).value, type(sheet.cell(i, col).value), sheet.name)
            caf_num = ''
            try:
                caf_num = try_to_str_int(sheet.cell(i, 0).value)
                payment_amt = float(sheet.cell(i, 1).value)
                # print caf_num, payment_amt, type(payment_amt)
                # get all objects related to given CAF number
                paid_user_obj_list = list(SupremeModel.objects.filter(caf_num=caf_num))
                # print paid_user_obj_list
                # Sort to get latest object first
                paid_user_obj_list.sort(key=lambda x: x.date_created, reverse=True)
                if paid_user_obj_list:
                    latest_user_obj = paid_user_obj_list[0]
                    remain_payment = float(latest_user_obj.pending_amt) - payment_amt
                    # print latest_user_obj.account_balance, payment_amt, remain_payment
                    if remain_payment < 100:
                        latest_user_obj.status = "Paid"
                    else:
                        latest_user_obj.status = "Partial Paid"
                    # print "old pending", float(latest_user_obj.pending_amt)
                    latest_user_obj.pending_amt = float(latest_user_obj.pending_amt) - payment_amt
                    # print "new pending", float(latest_user_obj.pending_amt)
                    latest_user_obj.save()
                    uploaded.append(caf_num)
                else:
                    not_found_caf.append(caf_num)
            except Exception, e:
                print traceback.format_exc()
                error_caf_no.append(caf_num)
    except Exception, e:
        print traceback.format_exc()
        return [{"error": "System Error: " + str(e)}]

    if uploaded:
        msg.append({"success": "Following entries Updated: {}".format(', '.join(uploaded))})
        u = UploadFileHistory()
        u.file_name = ufile
        u.sheet_no = sheet_no
        u.based_on = based_on
        u.speed = speed
        u.upload_type = "P"
        u.uploaded_date = datetime.datetime.now()
        u.no_of_records = len(uploaded)
        u.user = user
        u.save()
    else:
        msg.append({"warning": "No Entries Updated"})

    if not_found_caf:
        msg.append({"error": "Following entries Not found in Records: {}".format(', '.join(not_found_caf))})
        pass
    else:
        # msg.append("All Given Entries Updated")
        pass

    if error_caf_no:
        msg.append({"error": "Error in Following CAF Numbers: {}".format(', '.join(error_caf_no))})
        pass
    else:
        pass
        # msg.append("No Error in any Entries")

    return msg


def fast_upload_paid_excel_file(ufile, sheet_no, based_on, speed, user):
    print('fast reading paid file', ufile, speed)
    wb = xlrd.open_workbook(file_contents=ufile.read())
    # print wb.nsheets
    if sheet_no not in range(1, wb.nsheets + 1):
        return [{"error": "Sheet number {} not in range. There are {} sheets".format(sheet_no, wb.nsheets)}]

    sheet = wb.sheet_by_index(sheet_no - 1)
    msg = []
    uploaded = []
    not_found_caf = []
    error_caf_no = []
    error_pending_amt_caf = []
    error_caf_no_range = []
    try:
        for range_list in chunks(range(1, sheet.nrows), speed):
            # print(range_list)
            # for col in range(sheet.ncols):
            #     print(col, sheet.cell(i, col).value, type(sheet.cell(i, col)), sheet.name)
            model_list = []
            # caf_num_list = []
            # payment_amt_list = []
            caf_payment_dict = {}
            try:
                for i in range_list:
                    caf_payment_dict[try_to_str_int(sheet.cell(i, 0).value)] = float(sheet.cell(i, 1).value)
                    # caf_num_list.append(try_to_str_int(sheet.cell(i, 0).value))
                    # payment_amt_list.append(float(sheet.cell(i, 1).value))
                # print caf_num_list, payment_amt_list
                print "CAF", caf_payment_dict
                # try:
                # get all objects related to given CAF number
                paid_user_obj_list = list(SupremeModel.objects.filter(caf_num__in=caf_payment_dict.keys()))
                print paid_user_obj_list
                paid_user_obj_dict = {}
                if paid_user_obj_list:
                    for paid_user in paid_user_obj_list:
                        if paid_user.caf_num in paid_user_obj_dict.keys():
                            paid_user_obj_dict[paid_user.caf_num].append(paid_user)
                        else:
                            paid_user_obj_dict[paid_user.caf_num] = [paid_user]
                print paid_user_obj_dict
                # Sort to get latest object first
                # paid_user_obj_list.sort(key=lambda x: x.date_created, reverse=True)
                # print paid_user_obj_list

                if paid_user_obj_dict:
                    for caf_num, paid_user_obj_list in paid_user_obj_dict.items():
                        # latest_user_obj = paid_user_obj_list[0]
                        try:
                            print(paid_user_obj_list)
                            paid_user_obj_list.sort(key=lambda x: x.date_created, reverse=True)
                            latest_user_obj = paid_user_obj_list[0]
                            payment_amt = caf_payment_dict[latest_user_obj.caf_num]
                            print latest_user_obj, payment_amt, latest_user_obj.pending_amt
                            # if float(latest_user_obj.pending_amt) > payment_amt:

                            latest_user_obj.pending_amt = Decimal(
                                format(float(latest_user_obj.pending_amt) - payment_amt, '.2f'))
                            if latest_user_obj.pending_amt < 100:
                                latest_user_obj.status = "Paid"
                            else:
                                latest_user_obj.status = "Partial Paid"
                            print latest_user_obj.pending_amt, payment_amt, latest_user_obj.status
                            # print "old pending", float(latest_user_obj.pending_amt)
                            # latest_user_obj.pending_amt = float(latest_user_obj.pending_amt) - payment_amt
                            # print "new pending", float(latest_user_obj.pending_amt)
                            # latest_user_obj.save()
                            model_list.append(latest_user_obj)
                            uploaded.append(latest_user_obj.caf_num)

                            # else:
                            #     latest_user_obj.pending_amt = Decimal(format(float(0.00), '.2f'))
                            #     latest_user_obj.status = "Paid"

                            # if latest_user_obj.pending_amt < 100:
                            #     latest_user_obj.status = "Paid"
                            # else:
                            #     latest_user_obj.status = "Partial Paid"

                            # print "old pending", float(latest_user_obj.pending_amt)
                            # latest_user_obj.pending_amt = float(latest_user_obj.pending_amt) - payment_amt
                            # print "new pending", float(latest_user_obj.pending_amt)
                            # latest_user_obj.save()
                            # model_list.append(latest_user_obj)
                            # uploaded.append(latest_user_obj.caf_num)
                            # error_pending_amt_caf.append(latest_user_obj.caf_num)
                        except:
                            error_caf_no.append(caf_num)
                else:
                    pass
                    # not_found_caf.append(latest_user_obj.caf_num)
            except Exception as e:
                # print traceback.format_exc()
                error_caf_no += caf_payment_dict.keys()
            # except Exception as e:
            #     print traceback.format_exc()
            #     error_caf_no_range.append(
            #         str(sheet.cell(range_list[0], 0).value) + " : " + str(sheet.cell(range_list[-1], 0).value))
            # print model_list
            if model_list:
                SupremeModel.objects.bulk_update(model_list, update_fields=['status', 'pending_amt'])
    except Exception as e:
        print traceback.format_exc()
        return [{"error": "System Error: " + str(e)}]

    if uploaded:
        msg.append({"success": "Following entries Updated: {}".format(', '.join(uploaded))})
        u = UploadFileHistory()
        u.file_name = ufile
        u.sheet_no = sheet_no
        u.based_on = based_on
        u.speed = speed
        u.upload_type = "P"
        u.uploaded_date = datetime.datetime.now()
        u.no_of_records = len(uploaded)
        u.user = user
        u.save()
    else:
        msg.append({"warning": "No Entries Updated"})

    if not_found_caf:
        msg.append({"warning": "Following entries Not found in Records: {}".format(', '.join(not_found_caf))})
        pass
    else:
        # msg.append("All Given Entries Updated")
        pass

    if error_pending_amt_caf:
        msg.append({"error": "Following entries Payment amount is grater than Pending amount: {}".format(
            ', '.join(error_pending_amt_caf))})
        pass
    else:
        pass

    if error_caf_no:
        msg.append({"error": "Error in Following CAF Numbers: {}".format(', '.join(error_caf_no))})
        pass
    else:
        # msg.append("No Error in any Entries")
        pass

        # if error_caf_no_range:
        #     msg.append({"error": "Something Wrong in CAF range : {}".format(', '.join(error_caf_no_range))})
        #     pass
        # else:
        # msg.append("No Error in any range of Entries")
        # pass

    return msg


@login_required
def paid_upload(request):
    session_id = request.META['HTTP_COOKIE'].split()[1].split('=')[1]
    print session_id, "UPLOAD REQ"
    message_list = []
    print 'rm', request.method
    if request.method == 'POST':
        form = PaidUploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            ufile = request.FILES['file']
            sheel_no = int(form.data['sheet_no'])
            based_on = form.data['based_on']
            speed = int(form.data['speed'])
            if ufile.name.endswith('.xls') or ufile.name.endswith('.xlsx'):
                if based_on == "Normal Upload":
                    msg = upload_paid_excel_file(ufile, sheel_no, based_on, speed, request.user)
                else:
                    msg = fast_upload_paid_excel_file(ufile, sheel_no, based_on, speed, request.user)
            else:
                msg = [{"error": " .xls, .xlsx and .csv file formats are only supported."}]
            message_list = msg
        else:
            message_list.append({"error": "Supply appropriate data."})
    form = PaidUploadFileForm()
    print message_list
    for msg in message_list:
        for k, v in msg.iteritems():
            if k == "warning":
                messages.warning(request, v)
            elif k == "error":
                print v
                messages.error(request, v)
            elif k == "success":
                messages.success(request, v)

    results = UploadFileHistory.objects.filter(upload_type='P').order_by('-uploaded_date').values()
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


def create_temp_xlsx_data_file(supreme_app_data):
    """
    This will create temp xlsx file and return that path
    :return:
    """
    file_path = '/tmp/temp%s.xlsx' % (str(time.time()))
    book = xlsxwriter.Workbook(file_path)
    sheet = book.add_worksheet('DATA SHEET-1')
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
    date_time_format_bold = book.add_format({'num_format': 'mm/dd/yy hh:mm AM/PM', 'bold': True})
    date_time_format_green_bg = book.add_format(
        {'num_format': 'mm/dd/yy hh:mm AM/PM', 'bg_color': 'green', 'border': 1})
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
        sheet.write(0, i, h, bold)

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
        sheet.write(j, i_var.inc_var, data.final_calling_date, date_time_format_bold)
        sheet.write(j, i_var.inc_var, data.final_calling_code, bold)
        sheet.write(j, i_var.inc_var, data.final_calling_remarks, bold)
        sheet.write(j, i_var.inc_var, data.final_followup_date, date_time_format_bold)
        sheet.write(j, i_var.inc_var, data.status, bold)

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


def create_temp_xlsx_report_file(supreme_app_data):
    """
    This will create temp xlsx report file and return that path
    :return:
    """
    file_path = '/tmp/temp%s.xlsx' % (str(time.time()))
    book = xlsxwriter.Workbook(file_path)
    sheet = book.add_worksheet('REPORTS-1')
    # sheet.set_tab_color('red')

    # Different formats for attractive xls file
    bold = book.add_format({'bold': True})
    red_font = book.add_format({'bold': True, 'font_color': 'red'})
    red_bg = book.add_format({'bg_color': 'red', 'border': 1})
    yellow_bg = book.add_format({'bg_color': 'yellow', 'border': 1, 'bold': True})
    orange_bg = book.add_format({'bg_color': 'orange', 'border': 1})
    green_bg = book.add_format({'bg_color': 'green', 'border': 1})

    date_format = book.add_format({'num_format': 'mm/dd/yyyy'})
    time_format = book.add_format({'num_format': 'hh:mm:ss'})
    date_time_format = book.add_format({'num_format': 'mm/dd/yy hh:mm AM/PM'})
    date_time_format_bold = book.add_format({'num_format': 'mm/dd/yy hh:mm AM/PM', 'bold': True})
    date_time_format_green_bg = book.add_format(
        {'num_format': 'mm/dd/yy hh:mm AM/PM', 'bg_color': 'green', 'border': 1})
    percent_format = book.add_format({'num_format': '0.00%'})

    # Set the columns widths.
    sheet.set_column("A:G", 15)

    # *****************         Cluster wise Performance        *****************
    caption = "Cluster wise Performance"
    sheet.write('C2:F2', caption, bold)

    cluster_data = supreme_app_data.values('cluster').annotate(alloc_cnt=Sum('no_of_active_services'),
                                                               alloc_val=Sum('account_balance'))
    cluster_paid_data = supreme_app_data.filter(status='Paid').values('cluster').annotate(
        res_cnt=Sum('no_of_active_services'),
        res_val=Sum('account_balance'))

    # print cluster_data
    # print cluster_paid_data
    cluster_paid_data_dict = {i['cluster']: i for i in cluster_paid_data}
    # print cluster_paid_data_dict
    cluster_wise_performance_data = []
    total_cluster = ['Totals', ]
    for cdata in cluster_data:
        cluster_list = [cdata['cluster'], cdata['alloc_cnt'], cdata['alloc_val']]
        if cdata['cluster'] in cluster_paid_data_dict:
            cluster_list.append(0 if cluster_paid_data_dict[cdata['cluster']]['res_cnt'] is None else
                                cluster_paid_data_dict[cdata['cluster']]['res_cnt'])
            cluster_list.append(0 if cluster_paid_data_dict[cdata['cluster']]['res_val'] is None else
                                cluster_paid_data_dict[cdata['cluster']]['res_val'])
            cluster_list.append(float(cluster_list[3] / cluster_list[1]))
            cluster_list.append(float(cluster_list[4] / cluster_list[2]))
        else:
            cluster_list.append(0)
            cluster_list.append(0)
            cluster_list.append(0)
            cluster_list.append(0)
        cluster_wise_performance_data.append(cluster_list)
    # print cluster_wise_performance_data
    for k in range(1, 5):
        total_cluster.append(sum(i[k] for i in cluster_wise_performance_data))
    total_cluster.append(float(total_cluster[3] / total_cluster[1]))
    total_cluster.append(float(total_cluster[4] / total_cluster[2]))
    cluster_wise_performance_data.append(total_cluster)
    column_list = [
        {'header': 'Cluster',
         # 'total_string': 'Totals'
         },
        {'header': 'Alloc Cnt',
         # 'total_function': 'sum',
         },
        {'header': 'Alloc Val',
         # 'total_function': 'sum',
         },
        {'header': 'Res Cnt',
         # 'total_function': 'sum',
         },
        {'header': 'Res Val',
         # 'total_function': 'sum',
         },
        {'header': 'Res Cnt %',
         # 'total_function': 'average',
         'format': percent_format,
         },
        {'header': 'Res Val%',
         # 'total_function': 'average',
         'format': percent_format,
         },
    ]
    options = {
        'data': cluster_wise_performance_data,
        'style': 'Table Style Medium 9',
        'total_row': True,
        # 'autofilter': False,
        'banded_rows': False, 'banded_columns': True,
        'first_column': True, 'last_column': True,
        # 'name': 'Cluster wise Performance',
        'columns': column_list
    }

    last_raw = 3 + len(cluster_wise_performance_data)
    sheet.add_table(2, 0, last_raw, len(column_list) - 1, options)

    # *****************         TC wise Performance         *****************
    caption = "TC wise Performance"
    sheet.write('C' + str(last_raw + 2), caption, bold)

    tc_data = supreme_app_data.values('final_tc_name').annotate(alloc_cnt=Sum('no_of_active_services'),
                                                                alloc_val=Sum('account_balance'))
    tc_paid_data = supreme_app_data.values('final_tc_name').filter(status='Paid').annotate(
        res_cnt=Sum('no_of_active_services'),
        res_val=Sum('account_balance'))

    # print tc_data
    # print tc_paid_data
    tc_paid_data_dict = {i['final_tc_name']: i for i in tc_paid_data}
    # print tc_paid_data_dict
    tc_wise_performance_data = []
    total_tc = ['Totals', ]
    for tdata in tc_data:
        tc_list = [tdata['final_tc_name'], tdata['alloc_cnt'], tdata['alloc_val']]
        if tdata['final_tc_name'] in tc_paid_data_dict:
            tc_list.append(0 if tc_paid_data_dict[tdata['final_tc_name']]['res_cnt'] is None else
                           tc_paid_data_dict[tdata['final_tc_name']]['res_cnt'])
            tc_list.append(0 if tc_paid_data_dict[tdata['final_tc_name']]['res_val'] is None else
                           tc_paid_data_dict[tdata['final_tc_name']]['res_val'])
            tc_list.append(float(tc_list[3] / tc_list[1]))
            tc_list.append(float(tc_list[4] / tc_list[2]))
        else:
            tc_list.append(0)
            tc_list.append(0)
            tc_list.append(0)
            tc_list.append(0)
        tc_wise_performance_data.append(tc_list)
    # print tc_wise_performance_data
    for k in range(1, 5):
        total_tc.append(sum(i[k] for i in tc_wise_performance_data))
    total_tc.append(float(total_tc[3] / total_tc[1]))
    total_tc.append(float(total_tc[4] / total_tc[2]))
    tc_wise_performance_data.append(total_tc)
    column_list = [
        {'header': 'TC Name',
         # 'total_string': 'Totals'
         },
        {'header': 'Alloc Cnt',
         # 'total_function': 'sum',
         },
        {'header': 'Alloc Val',
         # 'total_function': 'sum',
         },
        {'header': 'Res Cnt',
         # 'total_function': 'sum',
         },
        {'header': 'Res Val',
         # 'total_function': 'sum',
         },
        {'header': 'Res Cnt %',
         # 'total_function': 'average',
         'format': percent_format,
         },
        {'header': 'Res Val%',
         # 'total_function': 'average',
         'format': percent_format,
         },
    ]
    options = {
        'data': tc_wise_performance_data,
        'style': 'Table Style Medium 10',
        'total_row': True,
        # 'autofilter': False,
        'banded_rows': False, 'banded_columns': True,
        'first_column': True, 'last_column': True,
        # 'name': 'Cluster wise Performance',
        'columns': column_list
    }
    nlast_raw = last_raw + len(tc_wise_performance_data) + 4
    sheet.add_table(last_raw + 3, 0, nlast_raw, len(column_list) - 1, options)

    # *****************         Product wise Performance         *****************
    caption = "Product wise Performance"
    sheet.write('C' + str(nlast_raw + 2), caption, bold)

    product_data = supreme_app_data.values('service_subtype').annotate(alloc_cnt=Sum('no_of_active_services'),
                                                                       alloc_val=Sum('account_balance'))
    product_paid_data = supreme_app_data.values('service_subtype').filter(status='Paid').annotate(
        res_cnt=Sum('no_of_active_services'),
        res_val=Sum('account_balance'))

    # print product_data
    # print product_paid_data
    product_paid_data_dict = {i['service_subtype']: i for i in product_paid_data}
    # print product_paid_data_dict
    product_wise_performance_data = []
    total_product = ['Totals', ]
    for tdata in product_data:
        product_list = [tdata['service_subtype'], tdata['alloc_cnt'], tdata['alloc_val']]
        if tdata['service_subtype'] in product_paid_data_dict:
            product_list.append(0 if product_paid_data_dict[tdata['service_subtype']]['res_cnt'] is None else
                                product_paid_data_dict[tdata['service_subtype']]['res_cnt'])
            product_list.append(0 if product_paid_data_dict[tdata['service_subtype']]['res_val'] is None else
                                product_paid_data_dict[tdata['service_subtype']]['res_val'])
            product_list.append(float(product_list[3] / product_list[1]))
            product_list.append(float(product_list[4] / product_list[2]))
        else:
            product_list.append(0)
            product_list.append(0)
            product_list.append(0)
            product_list.append(0)
        product_wise_performance_data.append(product_list)
    # print product_wise_performance_data
    for k in range(1, 5):
        total_product.append(sum(i[k] for i in product_wise_performance_data))
    total_product.append(float(total_product[3] / total_product[1]))
    total_product.append(float(total_product[4] / total_product[2]))
    product_wise_performance_data.append(total_product)
    column_list = [
        {'header': 'Product',
         # 'total_string': 'Totals'
         },
        {'header': 'Alloc Cnt',
         # 'total_function': 'sum',
         },
        {'header': 'Alloc Val',
         # 'total_function': 'sum',
         },
        {'header': 'Res Cnt',
         # 'total_function': 'sum',
         },
        {'header': 'Res Val',
         # 'total_function': 'sum',
         },
        {'header': 'Res Cnt %',
         # 'total_function': 'average',
         'format': percent_format,
         },
        {'header': 'Res Val%',
         # 'total_function': 'average',
         'format': percent_format,
         },
    ]
    options = {
        'data': product_wise_performance_data,
        'style': 'Table Style Medium 11',
        'total_row': True,
        # 'autofilter': False,
        'banded_rows': False, 'banded_columns': True,
        'first_column': True, 'last_column': True,
        # 'name': 'Cluster wise Performance',
        'columns': column_list
    }
    last_raw = nlast_raw + len(product_wise_performance_data) + 4
    sheet.add_table(nlast_raw + 3, 0, last_raw, len(column_list) - 1, options)

    # *****************         Day Wie Resolution Trend         *****************
    caption = "Day Wie Resolution Trend"
    sheet.write('C' + str(last_raw + 2), caption, bold)

    # print "\n\n\n\n"
    SM = supreme_app_data.values('cluster').annotate(alloc_cnt=Sum('no_of_active_services'))

    days = list(set(map(lambda x: x['date_modified'].date(), supreme_app_data.values('date_modified').distinct())))
    # print days
    days.sort()
    clusters = map(lambda x: x['cluster'], SupremeModel.objects.values('cluster').distinct())

    # print "\n\n"
    day_wise_trend_data = []
    day_wise_trend_percent = ['Per Day Res Cnt %', '']
    total_day_wise = ['Totals', ]
    for tdata in SM:
        day_tc_list = [tdata['cluster'], tdata['alloc_cnt']]
        # print "\n"
        for day in days:
            day_paid_data = supreme_app_data.filter(status='Paid', date_modified__year=day.strftime("%Y"),
                                                    date_modified__month=day.strftime("%m"),
                                                    date_modified__day=day.strftime("%d"),
                                                    cluster=tdata['cluster']).aggregate(
                res_cnt=Sum('no_of_active_services'))
            if day_paid_data['res_cnt']:
                day_tc_list.append(day_paid_data['res_cnt'])
                day_wise_trend_percent.append(float(day_paid_data['res_cnt'] / tdata['alloc_cnt']))
            else:
                day_tc_list.append(0)
                day_wise_trend_percent.append(0.0)
        day_wise_trend_data.append(day_tc_list)
    # print day_wise_trend_data
    total_day_wise.append(sum(i[1] for i in day_wise_trend_data))
    k = 2
    for day in days:
        total_day_wise.append(sum(i[k] for i in day_wise_trend_data))
        k += 1
    day_wise_trend_data.append(total_day_wise)

    column_list = [
        {'header': 'Product',
         # 'total_string': 'Totals'
         },
        {'header': 'Alloc Cnt',
         # 'total_function': 'sum',
         },
    ]
    for day in days:
        column_list.append({'header': day.strftime("%d/%B"),
                            # 'format': date_format,
                            # 'total_function': 'sum',
                            })
    # print column_list

    # for SMmodel in SM:
    #     print SupremeModel.objects.filter(status='Paid').filter(cluster=SMmodel['cluster']).values(
    #         'date_modified').annotate(
    #         alloc_cnt=Sum('no_of_active_services'))
    #     day_wise_trend_data[SMmodel['cluster']] = {}

    options = {
        'data': day_wise_trend_data,
        'style': 'Table Style Medium 12',
        'total_row': True,
        # 'autofilter': False,
        'banded_rows': False, 'banded_columns': True,
        'first_column': True, 'last_column': False,
        # 'name': 'Cluster wise Performance',
        'columns': column_list
    }

    nlast_raw = last_raw + len(day_wise_trend_data) + 4
    sheet.add_table(last_raw + 3, 0, nlast_raw, len(column_list) - 1, options)

    day_wise_trend_percent_tmp = []
    day_wise_trend_percent_tmp.append(day_wise_trend_percent)
    # print day_wise_trend_percent_tmp

    column_list = [
        {'header': 'Product'
         },
        {'header': 'Alloc Cnt'
         },
    ]
    for day in days:
        column_list.append({'header': '',
                            'format': percent_format,
                            })

    options = {
        'data': day_wise_trend_percent_tmp,
        'style': 'Table Style Medium 12',
        # 'autofilter': False,
        'banded_rows': False, 'banded_columns': True,
        'first_column': True, 'last_column': False,
        # 'name': 'Cluster wise Performance',
        'columns': column_list,
        'header_row': False
    }

    last_raw = nlast_raw + len(day_wise_trend_percent_tmp) + 1
    sheet.add_table(nlast_raw + 1, 0, last_raw, len(column_list) - 1, options)
    """
    nlast_raw += 2
    sheet.merge_range('A%d:B%d' % (nlast_raw, nlast_raw), 'Per Day Res Cnt %', bold)

    col= 2
    for percent in day_wise_trend_percent:
        sheet.write(nlast_raw-1, col, "%.2f" %percent+"%")
        col += 1
    """

    # *****************         Invoice wise Performance        *****************
    caption = "Invoice wise Performance"
    sheet.write('C' + str(last_raw + 2), caption, bold)

    cluster_data = supreme_app_data.values('cluster').annotate(alloc_cnt=Sum('no_of_active_services'),
                                                               alloc_val=Sum('account_balance'))
    cluster_paid_data = supreme_app_data.filter(status='Paid').values('cluster').annotate(
        res_cnt=Sum('no_of_active_services'),
        res_val=Sum('account_balance'))

    # print cluster_data
    # print cluster_paid_data
    cluster_paid_data_dict = {i['cluster']: i for i in cluster_paid_data}
    # print cluster_paid_data_dict
    cluster_wise_performance_data = []
    for cdata in cluster_data:
        cluster_list = [cdata['cluster'], cdata['alloc_cnt'], cdata['alloc_val']]
        if cdata['cluster'] in cluster_paid_data_dict:
            cluster_list.append(0 if cluster_paid_data_dict[cdata['cluster']]['res_cnt'] is None else
                                cluster_paid_data_dict[cdata['cluster']]['res_cnt'])
            cluster_list.append(0 if cluster_paid_data_dict[cdata['cluster']]['res_val'] is None else
                                cluster_paid_data_dict[cdata['cluster']]['res_val'])
            cluster_list.append(float(cluster_list[3] / cluster_list[1]))
            cluster_list.append(float(cluster_list[4] / cluster_list[2]))
        else:
            cluster_list.append(0)
            cluster_list.append(0)
            cluster_list.append(0)
            cluster_list.append(0)
        cluster_wise_performance_data.append(cluster_list)

        # no_of_invoice_raised
        for i in range(1, 4):
            cluster_invoice_raised_data = supreme_app_data.filter(no_of_invoice_raised=float(i),
                                                                  cluster=cdata['cluster']).aggregate(
                alloc_cnt=Sum('no_of_active_services'), alloc_val=Sum('account_balance'))
            cluster_invoice_raised_paid_data = supreme_app_data.filter(status='Paid', no_of_invoice_raised=float(i),
                                                                       cluster=cdata['cluster']).aggregate(
                res_cnt=Sum('no_of_active_services'), res_val=Sum('account_balance'))
            print cluster_invoice_raised_data
            print cluster_invoice_raised_paid_data
            # print cluster_invoice_raised_paid_data
            # cluster_invoice_raised_paid_data_dict = {i['cluster']: i for i in cluster_invoice_raised_paid_data}
            # print cluster_paid_data_dict
            cluster_list = [i]
            if cluster_invoice_raised_data['alloc_cnt'] is not None:
                cluster_list.append(cluster_invoice_raised_data['alloc_cnt'])
                cluster_list.append(cluster_invoice_raised_data['alloc_val'])
            else:
                cluster_list.append(0)
                cluster_list.append(0)
            if cluster_invoice_raised_paid_data['res_cnt'] is not None:
                cluster_list.append(cluster_invoice_raised_paid_data['res_cnt'])
                cluster_list.append(cluster_invoice_raised_paid_data['res_val'])
                cluster_list.append(float(cluster_list[3] / cluster_list[1]))
                cluster_list.append(float(cluster_list[4] / cluster_list[2]))
            else:
                cluster_list.append(0)
                cluster_list.append(0)
                cluster_list.append(0)
                cluster_list.append(0)
            cluster_wise_performance_data.append(cluster_list)

    total_cluster_data = supreme_app_data.aggregate(alloc_cnt=Sum('no_of_active_services'),
                                                    alloc_val=Sum('account_balance'))
    total_cluster_paid_data = supreme_app_data.filter(status='Paid').aggregate(res_cnt=Sum('no_of_active_services'),
                                                                               res_val=Sum('account_balance'))

    cluster_list = ['Totals']
    cluster_list.append(total_cluster_data['alloc_cnt'])
    cluster_list.append(total_cluster_data['alloc_val'])
    cluster_list.append(0 if total_cluster_paid_data['res_cnt'] is None else total_cluster_paid_data['res_cnt'])
    cluster_list.append(0 if total_cluster_paid_data['res_val'] is None else total_cluster_paid_data['res_val'])
    cluster_list.append(float(cluster_list[3] / cluster_list[1]))
    cluster_list.append(float(cluster_list[4] / cluster_list[2]))

    cluster_wise_performance_data.append(cluster_list)
    # print cluster_wise_performance_data

    column_list = [
        {'header': 'Cluster',
         },
        {'header': 'Alloc Cnt',
         },
        {'header': 'Alloc Val',
         },
        {'header': 'Res Cnt',
         },
        {'header': 'Res Val',
         },
        {'header': 'Res Cnt %',
         'format': percent_format,
         },
        {'header': 'Res Val%',
         'format': percent_format,
         },
    ]
    options = {
        'data': cluster_wise_performance_data,
        'style': 'Table Style Medium 9',
        'total_row': True,
        # 'autofilter': False,
        'banded_rows': True, 'banded_columns': True,
        'first_column': True, 'last_column': True,
        # 'name': 'Cluster wise Performance',
        'columns': column_list
    }

    nlast_raw = last_raw + len(cluster_wise_performance_data) + 4
    sheet.add_table(last_raw + 3, 0, nlast_raw, len(column_list) - 1, options)

    # *****************         User Wise Data         *****************
    users = User.objects.all()
    for user in users:
        # caption = str(user) + " Report"
        # sheet.write('C' + str(nlast_raw + 2), caption, bold)

        column_list = [
            {'header': ' ',
             },
        ]

        column_list1 = [
            {'header': ' ',
             },
        ]

        user_wise_trend_data = []
        no_call_list = ['No.of Calls']
        repeat = ['Repeat']
        fresh = ['Fresh']
        repeat_percent = ['Repeat %']
        fresh_percent = ['Fresh %']
        contact = ['Contact']
        no_contact = ['No Contact']
        contact_percent = ['Contact %']
        no_contact_percent = ['No Contact %']
        DISPO_LIST = ['DISPO LIST']
        CB = ['CB']
        RR = ['RR']
        OS = ['OS']
        SO = ['SO']
        CLMPD = ['CLMPD']
        WPD = ['WPD']
        PAID = ['PAID']
        PARTPAID = ['PARTPAID']
        BD = ['BD']
        BNR = ['BNR']
        NR = ['NR']
        RTP = ['RTP']
        PTP = ['PTP']
        CANCELLATION = ['CANCELLATION']
        SALES_ISSUE = ['SALES ISSUE']
        WAIVERS = ['WAIVERS']
        Others = ['Others']

        call_days = []
        for day in days:
            user_supreme_app_data = supreme_app_data.filter(processed=True, final_calling_date__year=day.strftime("%Y"),
                                                            final_calling_date__month=day.strftime("%m"),
                                                            final_calling_date__day=day.strftime("%d"),
                                                            final_tc_name=user)

            total_call = user_supreme_app_data.count()
            if not total_call:
                continue
            call_days.append(day)
            column_list.append({'header': day.strftime("%d/%B"), })
            column_list1.append({'header': day.strftime("%d/%B"),
                                 'format': percent_format,
                                 })
            no_call_list.append(total_call)

            total_repeat = 0
            total_fresh = 0
            for supreme_data in user_supreme_app_data:
                total_record_count = supreme_data.tcmodel_set.count()
                if total_record_count > 1:
                    total_repeat += 1
                else:
                    total_fresh += 1

            repeat.append(total_repeat)
            fresh.append(total_fresh)

            if total_call != 0 and total_repeat != 0:
                repeat_percent.append(float(100 * total_repeat / total_call) / 100)
            else:
                repeat_percent.append(0)

            if total_call != 0 and total_fresh != 0:
                fresh_percent.append(float(100 * total_fresh / total_call) / 100)
            else:
                fresh_percent.append(0)

            no_contact_list = ['CB', 'RR', 'SO', 'NR']

            total_contact = user_supreme_app_data.filter(~Q(final_calling_code__in=no_contact_list)).count()
            contact.append(total_contact)

            total_no_contact = user_supreme_app_data.filter(final_calling_code__in=no_contact_list).count()
            no_contact.append(total_no_contact)

            if total_call != 0 and total_contact != 0:
                contact_percent.append(float(100 * total_contact / total_call) / 100)
            else:
                contact_percent.append(0)

            if total_call != 0 and total_no_contact != 0:
                no_contact_percent.append(float(100 * total_no_contact / total_call) / 100)
            else:
                no_contact_percent.append(0)

            DISPO_LIST.append('')

            total_CB = user_supreme_app_data.filter(final_calling_code='CB').count()
            CB.append(total_CB)

            total_RR = user_supreme_app_data.filter(final_calling_code='RR').count()
            RR.append(total_RR)

            total_OS = user_supreme_app_data.filter(final_calling_code='OS').count()
            OS.append(total_OS)

            total_SO = user_supreme_app_data.filter(final_calling_code='SO').count()
            SO.append(total_SO)

            total_CLMPD = user_supreme_app_data.filter(final_calling_code='CLMPD').count()
            CLMPD.append(total_CLMPD)

            total_WPD = user_supreme_app_data.filter(final_calling_code='WPD').count()
            WPD.append(total_WPD)

            total_PAID = user_supreme_app_data.filter(final_calling_code='PAID').count()
            PAID.append(total_PAID)

            total_PARTPAID = user_supreme_app_data.filter(final_calling_code='PARTPAID').count()
            PARTPAID.append(total_PARTPAID)

            total_BD = user_supreme_app_data.filter(final_calling_code='BD').count()
            BD.append(total_BD)

            total_BNR = user_supreme_app_data.filter(final_calling_code='BNR').count()
            BNR.append(total_BNR)

            total_NR = user_supreme_app_data.filter(final_calling_code='NR').count()
            NR.append(total_NR)

            total_RTP = user_supreme_app_data.filter(final_calling_code='RTP').count()
            RTP.append(total_RTP)

            total_PTP = user_supreme_app_data.filter(final_calling_code='PTP').count()
            PTP.append(total_PTP)

            total_CANCELLATION = user_supreme_app_data.filter(final_calling_code='CANCELLATION').count()
            CANCELLATION.append(total_CANCELLATION)

            total_SALES_ISSUE = user_supreme_app_data.filter(final_calling_code='SALES_ISSUE').count()
            SALES_ISSUE.append(total_SALES_ISSUE)

            total_WAIVERS = user_supreme_app_data.filter(final_calling_code='WAIVERS').count()
            WAIVERS.append(total_WAIVERS)

            total_Others = user_supreme_app_data.filter(final_calling_code='Others').count()
            Others.append(total_Others)

        if call_days:
            caption = str(user) + " Report"
            sheet.write('C' + str(nlast_raw + 2), caption, bold)

            # Total Column
            # column_list.append({'header': 'Totals',
            #                 'formula': '=SUM(Table7[@[%s]:[%s]])' % (str(call_days[0].strftime("%d/%B")), str(call_days[-1].strftime("%d/%B"))),
            #                 'total_function': 'sum',
            #                 })

            column_list.append({'header': 'Totals'})
            DISPO_LIST.append('')
            CB.append(sum(map(int, CB[1:])))
            RR.append(sum(map(int, RR[1:])))
            OS.append(sum(map(int, OS[1:])))
            SO.append(sum(map(int, SO[1:])))
            CLMPD.append(sum(map(int, CLMPD[1:])))
            WPD.append(sum(map(int, WPD[1:])))
            PAID.append(sum(map(int, PAID[1:])))
            PARTPAID.append(sum(map(int, PARTPAID[1:])))
            BD.append(sum(map(int, BD[1:])))
            BNR.append(sum(map(int, BNR[1:])))
            NR.append(sum(map(int, NR[1:])))
            RTP.append(sum(map(int, RTP[1:])))
            PTP.append(sum(map(int, PTP[1:])))
            CANCELLATION.append(sum(map(int, CANCELLATION[1:])))
            SALES_ISSUE.append(sum(map(int, SALES_ISSUE[1:])))
            WAIVERS.append(sum(map(int, WAIVERS[1:])))
            Others.append(sum(map(int, Others[1:])))
            no_call_list.append(sum(map(int, no_call_list[1:])))
            repeat.append(sum(map(int, repeat[1:])))
            fresh.append(sum(map(int, fresh[1:])))
            contact.append(sum(map(int, contact[1:])))
            no_contact.append(sum(map(int, no_contact[1:])))

            user_wise_trend_data.append(DISPO_LIST)
            user_wise_trend_data.append(CB)
            user_wise_trend_data.append(RR)
            user_wise_trend_data.append(OS)
            user_wise_trend_data.append(SO)
            user_wise_trend_data.append(CLMPD)
            user_wise_trend_data.append(WPD)
            user_wise_trend_data.append(PAID)
            user_wise_trend_data.append(PARTPAID)
            user_wise_trend_data.append(BD)
            user_wise_trend_data.append(BNR)
            user_wise_trend_data.append(NR)
            user_wise_trend_data.append(RTP)
            user_wise_trend_data.append(PTP)
            user_wise_trend_data.append(CANCELLATION)
            user_wise_trend_data.append(SALES_ISSUE)
            user_wise_trend_data.append(WAIVERS)
            user_wise_trend_data.append(Others)
            user_wise_trend_data.append(no_call_list)
            user_wise_trend_data.append(repeat)
            user_wise_trend_data.append(fresh)
            user_wise_trend_data.append(contact)
            user_wise_trend_data.append(no_contact)

            print column_list
            options = {
                'data': user_wise_trend_data,
                'style': 'Table Style Medium 10',
                'total_row': False,
                # 'autofilter': False,
                'banded_rows': False, 'banded_columns': True,
                'first_column': True, 'last_column': False,
                # 'name': 'Cluster wise Performance',
                'columns': column_list
            }

            last_raw = nlast_raw + len(user_wise_trend_data) + 2
            sheet.add_table(nlast_raw + 3, 0, last_raw, len(column_list) - 1, options)

            nlast_raw = last_raw
        if call_days:
            column_list1.append({'header': 'Avg',
                                 # 'formula': '=AVERAGE(Table8[@[%s]:[%s]])' % (str(call_days[0].strftime("%d/%B")), str(call_days[-1].strftime("%d/%B"))),
                                 # 'total_function': 'average',
                                 'format': percent_format,
                                 })

            user_wise_trend_data = []
            print repeat_percent[1:]
            repeat_percent.append(float(sum(map(float, repeat_percent[1:]))) / len(repeat_percent[1:]))
            fresh_percent.append(float(sum(map(float, fresh_percent[1:]))) / len(fresh_percent[1:]))
            contact_percent.append(float(sum(map(float, contact_percent[1:]))) / len(contact_percent[1:]))
            no_contact_percent.append(float(sum(map(float, no_contact_percent[1:]))) / len(no_contact_percent[1:]))
            user_wise_trend_data.append(repeat_percent)
            user_wise_trend_data.append(fresh_percent)
            user_wise_trend_data.append(contact_percent)
            user_wise_trend_data.append(no_contact_percent)

            options = {
                'data': user_wise_trend_data,
                'style': 'Table Style Medium 10',
                'total_row': 0,
                # 'autofilter': False,
                'banded_rows': False, 'banded_columns': True,
                'first_column': True, 'last_column': False,
                # 'name': 'Cluster wise Performance',
                'columns': column_list1,
                'header_row': False

            }

            last_raw = nlast_raw + len(user_wise_trend_data) + 2
            sheet.add_table(nlast_raw + 1, 0, last_raw, len(column_list1) - 1, options)

            nlast_raw = last_raw

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
            return render(request, 'SupremeApp/download.html', locals())
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
                path = create_temp_xlsx_data_file(supreme_app_data)
                response = HttpResponse(file(path, 'r').read())
                response['Content-Disposition'] = 'attachment;filename=SUPREME_DATA_from_{}_to_{}.xlsx'.format(
                    from_date, to_date)
                response['Content-Length'] = os.path.getsize(path)
                print path
                os.remove(path)
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
    print session_id, "Report Download REQ"
    messages = []
    if request.method == 'POST':
        print "POST REQ"
        form = RDownloadFileForm(request.POST)
        if not form.is_valid():
            settings.FORM_SESSION[session_id] = form
            return render(request, 'SupremeApp/report_download.html', locals())
        else:
            from_date = form.cleaned_data['from_date']
            to_date = form.cleaned_data['to_date']
            # based_on = form.data['based_on']
            print from_date, to_date
            supreme_app_data = get_supreme_app_data(datetime.datetime.combine(from_date, datetime.time(0, 0)),
                                                    datetime.datetime.combine(to_date, datetime.time(23, 59)),
                                                    based_on="Last Modified")
            if not supreme_app_data:
                messages.append(" No Search Results")
            else:
                path = create_temp_xlsx_report_file(supreme_app_data)
                response = HttpResponse(file(path, 'r').read())
                response['Content-Disposition'] = 'attachment;filename=SUPREME_DATA_from_{}_to_{}.xlsx'.format(
                    from_date, to_date)
                response['Content-Length'] = os.path.getsize(path)
                print path
                os.remove(path)
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
    if request.user.is_superuser:
        print "INDEX REQ", request
        return render(request, 'SupremeApp/index.html', locals())
    else:
        return HttpResponseRedirect("/SupremeApp/data/")


@login_required
def mainindex(request):
    if request.user.is_superuser:
        return render(request, 'index.html', locals())
    else:
        print "123"
        return HttpResponseRedirect("/SupremeApp/data/")


@login_required
def data(request):
    print 'data req', request
    return HttpResponseRedirect("/admin/SupremeApp/suprememodel/")
