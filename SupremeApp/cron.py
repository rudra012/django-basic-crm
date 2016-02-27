from string import lower
from threading import Timer

import os
import cronjobs
from django.core.mail import EmailMessage
from .models import *
from .views import get_supreme_app_data, create_temp_xlsx_report_file


@cronjobs.register()
def run_backup_mail_cron():
    # crontab -e
    # * * * * * /home/lintel/Mayur/url_design/django1.8env/bin/python /home/lintel/Mayur/url_design/Suprem2/supreme/manage.py cron run_backup_mail_cron
    print "In Cron Jobs"
    now = datetime.datetime.now()
    for st in Setting.objects.all():
        if lower(now.strftime("%a")) in st.days:
            print "correct"

            def send_backup():
                # date_1 = datetime.datetime.strptime(datetime.datetime.strftime(datetime.datetime.now(), '%m/%d/%y'), "%m/%d/%y")
                # from_date = datetime.datetime.strftime(date_1 - datetime.timedelta(days=int(st.no_of_backup_days)), '%Y-%m-%d 00:00:00')
                upload_history = UploadFileHistory().objects.filter(upload_type="D", generate_new_report=True).order_by("-uploaded_date")
                if upload_history:
                    from_date = upload_history[0].uploaded_date
                    to_date = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d 23:59:59')

                    file_name = 'SUPREME_DATA_from_%s_to_%s.xlsx' %(datetime.datetime.strftime(from_date, '%d-%m-%Y'), datetime.datetime.strftime(datetime.datetime.now(), '%d-%m-%Y'))

                    sender_list = []
                    subject = 'Daily Report File'
                    msg = file_name[:-5]
                    from_email = 'lintelservice001@gmail.com'

                    for ud in UserDetail.objects.all():
                        sender_list.append(ud.email)

                    mail = EmailMessage(subject, msg, from_email, sender_list)

                    # Attach Reoprt
                    print st.no_of_backup_days
                    print from_date, to_date

                    supreme_app_data = get_supreme_app_data(from_date, to_date, based_on="Last Modified")
                    # print supreme_app_data
                    if not supreme_app_data:
                        html_content = '<p>There is no backup for today.</p>'
                        mail.attach_alternative(html_content, "text/html")
                    else:
                        path = create_temp_xlsx_report_file(supreme_app_data)
                        print path
                        new_path = '/tmp/' + file_name
                        os.rename(path, new_path)
                        mail.attach_file(new_path)
                        os.remove(new_path)
                    try:
                        mail.send()
                    except Exception as err:
                        print err
                    print "Msg Sent"

            def get_sec(s):
                l = s.split(':')
                return float(l[0]) * 3600 + int(l[1]) * 60 + int(l[2])

            seconds = get_sec(str(st.time))
            print seconds

            t = Timer(seconds, send_backup)
            t.start()

        else:
            print "sorry"
