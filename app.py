import configparser
from flask import Flask, render_template, session, redirect, url_for, session, flash
from flask_wtf import FlaskForm
from wtforms import (StringField,
                     RadioField,SelectField,
                     SubmitField)
from wtforms.fields.html5 import EmailField
from wtforms.validators import DataRequired, Email
from backend import CreatePDF, PrepareInfo, Mailer, XLS_Writer

config = configparser.ConfigParser(default_section=None)
config.read('conf/config.conf', encoding='utf-8')

app = Flask(__name__)
app.config['SECRET_KEY'] = 'mysecretkey'


class ApplyForm(FlaskForm):

    lang = RadioField(choices=[('en', 'English'),
                               ('de', 'Deutsch')],
                                default='en')
    email = EmailField(validators=[DataRequired(),
                                   Email()])
    gender = RadioField(choices=[('not_known', 'Not Known'),
                                 ('male', 'Male'),
                                 ('female', 'Female')],
                                default='not_known')
    hr_person_name = StringField()
    company_name = StringField(validators=[DataRequired()])
    company_address = StringField(validators=[DataRequired()])
    source = SelectField(choices=config.items('job_sites'))
    other_source = StringField('Other Source')
    submit = SubmitField('Apply!')


@app.route('/', methods=['GET', 'POST'])
def index():
    form = ApplyForm()
    if form.validate_on_submit():
        session['lang'] = form.lang.data
        session['email'] = form.email.data
        session['gender'] = form.gender.data
        session['hr_person_name'] = form.hr_person_name.data
        session['company_name'] = form.company_name.data
        session['company_address'] = form.company_address.data
        session['source'] = form.source.data
        session['other_source'] = form.other_source.data

        # PREPARE TEXTs
        input = PrepareInfo(session)
        # GENERATE PDFs
        pdf = CreatePDF()
        cover = pdf.generate_cover(input)
        # SEND EMAIL
        if cover:
            mail = Mailer()
            sent = mail.send([session['email']], input)
        if sent:
            xls = XLS_Writer()
            xls.update(input)
        # REDIRECT
        return redirect(url_for("success"))

    return render_template('index.html', form=form)

@app.route('/success')
def success():

    return render_template('success.html')


if __name__ == '__main__':
    app.run()
