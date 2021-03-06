from django.db import models


# Create your models here.
class SupremeModel(models.Model):
    # Import Data
    caf_num = models.CharField(max_length=30, null=True, verbose_name="A/C No")
    cust_name = models.CharField(max_length=100, null=True, verbose_name="Customer Name")
    customer_category = models.CharField(max_length=50, null=True)
    risk_class_code = models.CharField(max_length=30, null=True)
    customer_type = models.CharField(max_length=50, null=True)
    customer_ecl = models.CharField(max_length=20, null=True)
    mdn_no = models.CharField(max_length=15, null=True, verbose_name="Mobile No.")
    adc_status = models.CharField(max_length=30, null=True)
    service_type = models.CharField(max_length=30, null=True)
    rate_plan = models.CharField(max_length=50, null=True, verbose_name="Rate Plan")
    otaf_date = models.DateField(null=True, verbose_name="Activation Date")
    phongen_status = models.CharField(max_length=30, null=True)
    desired_service = models.CharField(max_length=30, null=True)
    og_bar = models.CharField(max_length=30, null=True)
    account_balance = models.CharField(max_length=30, null=True, verbose_name="O/S Amt")
    pending_amount = models.CharField(max_length=30, null=True)
    debtors_age = models.CharField(max_length=20, null=True)
    voluntary_deposit = models.CharField(max_length=30, null=True)
    unbilled_ild = models.CharField(max_length=30, null=True)
    total_unbilled = models.CharField(max_length=20, null=True)
    billed_not_due_amount = models.CharField(max_length=20, null=True)
    b1 = models.CharField(max_length=30, null=True)
    b2 = models.CharField(max_length=30, null=True)
    b3 = models.CharField(max_length=30, null=True)
    b4 = models.CharField(max_length=30, null=True)
    b5 = models.CharField(max_length=30, null=True)
    b6 = models.CharField(max_length=30, null=True)
    b7 = models.CharField(max_length=30, null=True)
    b8 = models.CharField(max_length=30, null=True)
    billed_and_overdue_amount = models.CharField(max_length=30, null=True)
    billed_outstanding_amount = models.CharField(max_length=30, null=True)
    exposure = models.CharField(max_length=30, null=True)
    latest_bill_due_date = models.DateField(null=True)
    no_of_invoice_raised = models.CharField(max_length=20, null=True)
    total_invoice_amout = models.CharField(max_length=30, null=True)
    no_of_payments_made = models.CharField(max_length=30, null=True, verbose_name="No. of Payments Made")
    payments_made_till_date = models.CharField(max_length=30, null=True)
    total_adjustment_amount = models.CharField(max_length=30, null=True)
    dapo_amount = models.CharField(max_length=30, null=True)
    ciou_terr = models.CharField(max_length=30, null=True)
    zone_terr = models.CharField(max_length=30, null=True)
    cluster_terr = models.CharField(max_length=30, null=True)
    circle_terr = models.CharField(max_length=30, null=True)
    country_terr = models.CharField(max_length=30, null=True)
    deposit = models.CharField(max_length=30, null=True)
    no_of_active_services = models.CharField(max_length=30, null=True)
    ecs = models.CharField(max_length=30, null=True)
    daf_num = models.CharField(max_length=30, null=True)
    ciou_name = models.CharField(max_length=50, null=True)
    zone_name = models.CharField(max_length=50, null=True)
    cluster_name = models.CharField(max_length=50, null=True)
    circle_name = models.CharField(max_length=50, null=True)
    line_1 = models.CharField(max_length=200, null=True, verbose_name="Address 1")
    line_2 = models.CharField(max_length=200, null=True, verbose_name="Address 2")
    city = models.CharField(max_length=50, null=True, verbose_name="City")
    cluster = models.CharField(max_length=30, null=True, verbose_name="Cluster")
    state = models.CharField(max_length=30, null=True)
    country = models.CharField(max_length=30, null=True)
    postcode = models.CharField(max_length=30, null=True)
    last_payment_date = models.DateField(null=True)
    excltype = models.CharField(max_length=50, null=True)
    child = models.CharField(max_length=20, null=True)
    do_not_disturb = models.CharField(max_length=20, null=True)
    handset_model = models.CharField(max_length=30, null=True)
    oldest_open_inv_due_date = models.DateField(null=True)
    last_bill_issue_date = models.DateField(null=True)
    last_biil_amount = models.CharField(max_length=30, null=True)
    last_payment_amount = models.CharField(max_length=30, null=True)
    no_buf_debt_age = models.CharField(max_length=30, null=True)
    data_unbilled_amount = models.CharField(max_length=30, null=True)
    rconnect_plan = models.CharField(max_length=50, null=True)
    bb_plan = models.CharField(max_length=30, null=True)
    premier_category = models.CharField(max_length=30, null=True)
    service_subtype = models.CharField(max_length=30, null=True)
    last_month_score = models.CharField(max_length=20, null=True)
    last2_month_score = models.CharField(max_length=20, null=True)
    last3_month_score = models.CharField(max_length=20, null=True)
    last4_month_score = models.CharField(max_length=20, null=True)
    last5_month_score = models.CharField(max_length=20, null=True)
    last6_month_score = models.CharField(max_length=20, null=True)
    curr_payment_behaviour_class = models.CharField(max_length=30, null=True)
    prev_payment_behaviour_class = models.CharField(max_length=30, null=True)
    alternate_landline_number = models.CharField(max_length=20, null=True, verbose_name="Alt No. 1")
    alternate_mobile_number = models.CharField(max_length=20, null=True, verbose_name="Alt No. 2")
    email_id = models.CharField(max_length=50, null=True, verbose_name="Email ID")
    corp_cust_category = models.CharField(max_length=50, null=True)
    mnp_flag = models.CharField(max_length=30, null=True)
    bill_cycle = models.CharField(max_length=10, null=True, verbose_name="Bill Cycle")
    bill_delivery_mode = models.CharField(max_length=30, null=True, verbose_name="Bill Delivery Mode")
    company_name = models.CharField(max_length=70, null=True)
    barring_reason = models.CharField(max_length=30, null=True)
    barring_date = models.DateField(null=True)
    allocation_date = models.DateField(null=True)
    closing_date = models.DateField(null=True)
    address = models.CharField(max_length=300, null=True)

    tc_name = models.CharField(max_length=50, null=True)
    calling_date = models.DateTimeField(null=True, verbose_name="Calling Date")
    calling_code = models.CharField(max_length=100, null=True, choices=[("CB",) * 2,
                                                                        ("RR",) * 2,
                                                                        ("OS",) * 2,
                                                                        ("SO",) * 2,
                                                                        ("CLMPD",) * 2,
                                                                        ("WPD",) * 2,
                                                                        ("PAID",) * 2,
                                                                        ("PARTPAID",) * 2,
                                                                        ("BD",) * 2,
                                                                        ("BNR",) * 2,
                                                                        ("NR",) * 2,
                                                                        ("RTP",) * 2,

                                                                        ("PTP",) * 2,
                                                                        ("Others",) * 2, ],
                                    verbose_name="Dispo Code")
    calling_remarks = models.CharField(max_length=250, null=True, verbose_name="Remarks")
    followup_date = models.DateTimeField(null=True, verbose_name="CB/PTP Date", blank=True)

    final_tc_name = models.CharField(max_length=50, null=True)
    final_calling_date = models.DateTimeField(null=True, verbose_name="Final Calling Date")
    final_calling_code = models.CharField(max_length=100, null=True, verbose_name="Final Dispo Code")
    final_calling_remarks = models.CharField(max_length=250, null=True, verbose_name="Final Remarks")
    final_followup_date = models.DateTimeField(null=True, verbose_name="Final CB/PTP Date", blank=True)

    # Export Data
    tc_1_name = models.CharField(max_length=50, null=True)
    tc_1_attempt_date = models.DateTimeField(null=True)
    tc_1_attempt_code = models.CharField(max_length=100, null=True)
    tc_1_attempt_remarks = models.CharField(max_length=250, null=True)
    tc_1_attempt_followup = models.DateTimeField(null=True, verbose_name="1st CB/PTP Date", blank=True)

    tc_2_name = models.CharField(max_length=50, null=True)
    tc_2_attempt_date = models.DateTimeField(null=True)
    tc_2_attempt_code = models.CharField(max_length=100, null=True)
    tc_2_attempt_remarks = models.CharField(max_length=250, null=True)
    tc_2_attempt_followup = models.DateTimeField(null=True, verbose_name="2st CB/PTP Date", blank=True)

    tc_3_name = models.CharField(max_length=50, null=True)
    tc_3_attempt_date = models.DateTimeField(null=True)
    tc_3_attempt_code = models.CharField(max_length=100, null=True)
    tc_3_attempt_remarks = models.CharField(max_length=250, null=True)
    tc_3_attempt_followup = models.DateTimeField(null=True, verbose_name="3st CB/PTP Date", blank=True)

    tc_4_name = models.CharField(max_length=50, null=True)
    tc_4_attempt_date = models.DateTimeField(null=True)
    tc_4_attempt_code = models.CharField(max_length=100, null=True)
    tc_4_attempt_remarks = models.CharField(max_length=250, null=True)
    tc_4_attempt_followup = models.DateTimeField(null=True, verbose_name="4st CB/PTP Date", blank=True)

    tc_5_name = models.CharField(max_length=50, null=True)
    tc_5_attempt_date = models.DateTimeField(null=True)
    tc_5_attempt_code = models.CharField(max_length=100, null=True)
    tc_5_attempt_remarks = models.CharField(max_length=250, null=True)
    tc_5_attempt_followup = models.DateTimeField(null=True, verbose_name="5st CB/PTP Date", blank=True)

    tc_6_name = models.CharField(max_length=50, null=True)
    tc_6_attempt_date = models.DateTimeField(null=True)
    tc_6_attempt_code = models.CharField(max_length=100, null=True)
    tc_6_attempt_remarks = models.CharField(max_length=250, null=True)
    tc_6_attempt_followup = models.DateTimeField(null=True, verbose_name="6st CB/PTP Date", blank=True)

    status = models.CharField(max_length=100, null=True, default="Unpaid")

    processed = models.BooleanField(default=False)
    attempt = models.CharField(max_length=5, null=True, blank=True, default=0)
    date_created = models.DateTimeField(auto_now_add=True, auto_now=False, verbose_name="Data Uploaded")
    date_modified = models.DateTimeField(auto_now_add=False, auto_now=True)  #

    def __repr__(self):
        return "--".join([str(self.attempt), str(self.final_followup_date), self.cust_name])
