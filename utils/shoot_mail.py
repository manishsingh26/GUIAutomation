
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart


class StatusMail(object):

    def __init__(self, email_config_path):
        with open(email_config_path) as config:
            self.email_config = json.load(config)

        self.smtp_host = self.email_config["smtp_details"]["host"]
        self.mail_from = "manish.6.singh.ext@nokia.com"
        self.mail_to = "manish.6.singh.ext@nokia.com"
        self.mail_cc = []
        self.mail_initial = self.email_config["email_details"]["mail_initial"]
        self.mail_greeting = self.email_config["email_details"]["mail_greeting"]
        self.mail_signature_header = self.email_config["email_details"]["mail_signature_header"]
        self.mail_signature = self.email_config["email_details"]["mail_signature"]
        self.mail_footer = self.email_config["email_details"]["mail_footer"]

    def shoot_mail(self, image_path, status):
        file_path = image_path
        file_status = status

        mail_subject = self.email_config["email_details"]["mail_subject"].replace("task_status", file_status)
        mail_body = self.email_config["email_details"]["main_body"].replace("task_status", file_status)

        msg_root = MIMEMultipart('related')
        msg_root['Subject'] = mail_subject
        msg_root['From'] = self.mail_from
        msg_root['To'] = self.mail_to

        msg_root.preamble = 'This is a multi-part message in MIME format.'
        msg_alternative = MIMEMultipart('alternative')
        msg_root.attach(msg_alternative)

        html_body = self.mail_initial + "<br>" + "<br>" + self.mail_greeting + "<br>" + "<br>" + mail_body + "<br>" + "<br>" + "<br>"
        html_footer = "<br>" + "<br>" + self.mail_signature_header + "<br>" + self.mail_signature + "<br>" + "<br>" + self.mail_footer
        msg_text = MIMEText(html_body + '<br><img src="cid:image1">' + html_footer, 'html')
        msg_alternative.attach(msg_text)

        fp = open(file_path, 'rb')
        msg_image = MIMEImage(fp.read())
        fp.close()

        msg_image.add_header('Content-ID', '<image1>')
        msg_root.attach(msg_image)

        smtp = smtplib.SMTP(self.smtp_host)
        mail_receipt = [self.mail_to] + self.mail_cc
        smtp.sendmail(self.mail_from, mail_receipt, msg_root.as_string())
        smtp.quit()


# if __name__ == "__main__":
#
#     images_path_ = r"C:\Users\m4singh\PycharmProjects\GraphAutomationMRTG\ErrorDetection&Mail\InputImages"
#     shoot_mail(images_path_)
