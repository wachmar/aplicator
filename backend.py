import os
import random
import comtypes.client
import pythoncom
import configparser
from docxtpl import DocxTemplate
from pathlib import Path

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.utils import formataddr
from email import encoders

from openpyxl import load_workbook
from datetime import date

config = configparser.ConfigParser()
config.read('conf/config.conf', encoding='utf-8')


class PrepareInfo():

    def __init__(self, session):

        self.lang = session['lang']
        self.email = session["email"]
        self.company_name = session["company_name"]
        self.company_address = session["company_address"]
        self.source = session['source']
        self.other_source = session["other_source"]
        self.job_site = self.other_source \
            if self.source == 99 \
            else config["job_sites"][self.source]
        self.position = config["texts"]["position_"+self.lang]
        self.subject = config["texts"]["email_subject_"+self.lang] + " " + self.position
        self.gender = session['gender']
        self.hr_person_name = session['hr_person_name']
        self.name = self._compute_name(config["texts"]["mr_"+self.lang],
                                       config["texts"]["mrs_"+self.lang],
                                       config["texts"]["unknown_name_"+self.lang])
        if self.lang == "de" and self.gender == "male":
            self.mannlich = "r"
        else:
            self.mannlich = ""

    def _compute_name(self, male, female, other):

        if self.gender == "male" and self.hr_person_name is not None:
            if self.lang == "de":
                return male + " " + self.hr_person_name
            else:
                return male + "" + self.hr_person_name
        elif self.gender == "female" and self.hr_person_name is not None:
            if self.lang == "de":
                return female + " " + self.hr_person_name
            else:
                return female + "" + self.hr_person_name
        else:
            return other


class CreatePDF():

    def generate_cover(self, input):

        # CFG
        in_file_path = "files/cover/cover-"+input.lang+".docx"
        temp_file_path = "files/cover/" + str(random.randint(0, 50)) + ".docx"
        out_file_path = "files/cover/cover-"+input.lang+".pdf"

        # Remove old pdf file (not necessary)
        if os.path.isfile(Path(out_file_path)):
            os.remove(Path(out_file_path))

        # Fill in text
        data_to_fill = {'Recipient_name': input.name,
                        'Company_name': input.company_name,
                        'Company_address': input.company_address,
                        'Job_site': input.job_site,
                        'Position': input.position,
                        'Mannlich': input.mannlich
                        }

        template = DocxTemplate(Path(in_file_path))
        template.render(data_to_fill)
        template.save(Path(temp_file_path))

        # Convert to PDF
        wdFormatPDF = 17

        in_file = os.path.abspath(Path(temp_file_path))
        out_file = os.path.abspath(Path(out_file_path))

        word = comtypes.client.CreateObject('Word.Application',pythoncom.CoInitialize())
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()

        # Get rid of the temp file
        os.remove(Path(temp_file_path))
        return True


class Mailer:

    def __init__(self):

        self.port = config['mailing']['port']
        self.smtp_server_domain_name = config['mailing']['smtp_server_domain_name']
        self.sender_mail = config['mailing']['sender_mail']
        self.sender_name = config['mailing']['sender_name']
        self.password = config['mailing']['password']
        # self.bcc_mail = config['mailing']['bcc_mail']

    def send(self, emails, input):

        service = smtplib.SMTP(self.smtp_server_domain_name, self.port)
        service.starttls()
        service.login(self.sender_mail, self.password)

        for email in emails:
            mail = MIMEMultipart('alternative')
            mail['Subject'] = input.subject
            mail['From'] = formataddr((self.sender_name, self.sender_mail))
            mail['To'] = input.email
            # mail['Bcc'] = self.bcc_mail

            f = open("files/email/email-"+input.lang+".txt", "r")
            text_template = f.read()
            f.close()
            f = open("files/email/email-"+input.lang+".html", "r")
            html_template = f.read()
            f.close()

            html_content = MIMEText(html_template.format(input.mannlich,
                                                         input.name,
                                                         input.job_site,
                                                         input.position,
                                                         ), 'html')
            text_content = MIMEText(text_template.format(input.mannlich,
                                                         input.name,
                                                         input.job_site,
                                                         input.position
                                                         ), 'plain')

            mail.attach(text_content)
            mail.attach(html_content)

            # attachment
            for file in ("cv", "cover"):
                file_path = f"files/{file}/{file}-{input.lang}.pdf"
                mimeBase = MIMEBase("application", "octet-stream")
                with open(file_path, "rb") as file:
                    mimeBase.set_payload(file.read())
                encoders.encode_base64(mimeBase)
                mimeBase.add_header("Content-Disposition", f"attachment; filename={Path(file_path).name}")
                mail.attach(mimeBase)

            # send
            service.sendmail(self.sender_mail, email, mail.as_string())

        # close connection
        service.quit()
        return True


class XLS_Writer():

    def update(self, input):
        xls_path = "files/applied.xlsx"
        wb = load_workbook(xls_path)
        page = wb.active

        info = [input.company_name,
                input.email,
                input.hr_person_name,
                input.lang,
                date.today().strftime("%d/%m/%Y"),
                input.job_site]
        page.append(info)

        wb.save(filename=xls_path)
