from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SubmitField, FileField, SelectField, FloatField, IntegerField
from wtforms.validators import DataRequired, Email, Length, EqualTo, Optional

class LoginForm(FlaskForm):
    email = StringField("Email", validators=[DataRequired(), Email()])
    password = PasswordField("Kata Sandi", validators=[DataRequired()])
    submit = SubmitField("Login")

class RegisterForm(FlaskForm):
    name = StringField("Nama", validators=[DataRequired(), Length(min=2, max=120)])
    email = StringField("Email", validators=[DataRequired(), Email()])
    password = PasswordField("Kata Sandi", validators=[DataRequired(), Length(min=6)])
    confirm = PasswordField("Konfirmasi Kata Sandi", validators=[DataRequired(), EqualTo("password")])
    submit = SubmitField("Daftar")

class UploadForm(FlaskForm):
    file = FileField("Upload Excel (.xlsx)", validators=[DataRequired()])
    submit = SubmitField("Upload")

class CriteriaForm(FlaskForm):
    name = StringField("Nama Kriteria", validators=[DataRequired()])
    ctype = SelectField("Tipe", choices=[("benefit","benefit"),("cost","cost")], validators=[DataRequired()])
    display_order = IntegerField("Urutan (ROC)", validators=[DataRequired()])
    submit = SubmitField("Simpan")

class SubcriteriaForm(FlaskForm):
    criteria_id = SelectField("Kriteria", coerce=int, validators=[DataRequired()])
    name = StringField("Nama Subkriteria", validators=[DataRequired()])
    min_val = FloatField("Min", validators=[Optional()])
    submit = SubmitField("Simpan")

