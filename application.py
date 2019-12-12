import os

from flask import (
    Flask,
    render_template,
    request
)
from flask_mail import (
    Mail,
    Message
)
from flask_session import Session
from tempfile import mkdtemp
from werkzeug import secure_filename
from werkzeug.exceptions import (
    default_exceptions,
    HTTPException,
    InternalServerError
)

from helpers import (
    connect_database,
    create_database,
    excel_convert,
    generate_filename
)

# Configure application
app = Flask(__name__, instance_relative_config=True)

os.makedirs(os.path.join(app.instance_path, 'htmlfi'), exist_ok=True)
os.makedirs(os.path.join(app.instance_path, 'output'), exist_ok=True)

# Ensure templates are auto-reloaded
app.config["TEMPLATES_AUTO_RELOAD"] = True

# Ensure responses aren't cached
@app.after_request
def after_request(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Expires"] = 0
    response.headers["Pragma"] = "no-cache"
    return response


# Configure session to use filesystem (instead of signed cookies)
app.config["SESSION_FILE_DIR"] = mkdtemp()
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)


@app.route("/", methods=["GET"])
def index():
    """Homepage for Forms-to-Data"""

    return render_template("index.html")


@app.route("/convert", methods=["GET", "POST"])
def history():
    """The bread and butter. The main event."""

    if request.method == "POST":
        # Set email settings
        email_settings = {
            "MAIL_SERVER": "smtp.gmail.com",
            "MAIL_PORT": 465,
            "MAIL_USE_TLS": False,
            "MAIL_USE_SSL": True,
            "MAIL_USERNAME": os.environ['EMAIL_USER'],
            "MAIL_PASSWORD": os.environ['EMAIL_PASSWORD']
        }

        app.config.update(email_settings)
        mail = Mail(app)

        # Get input from the posted form
        user_filename = generate_filename()

        xlsxFile = request.files['xlsxFile']
        xlsx_filename = secure_filename(xlsxFile.filename)
        xlsxFile.save(os.path.join(app.instance_path, 'htmlfi', xlsx_filename))
        xl = os.path.join(app.instance_path, 'htmlfi', xlsx_filename)

        user_query = request.files["queryList"]
        q_filename = secure_filename(user_query.filename)
        user_query.save(os.path.join(app.instance_path, 'htmlfi', q_filename))
        query = os.path.join(app.instance_path, 'htmlfi', q_filename)

        user_email = request.form.get("userEmail")

        converted_csv = excel_convert(xl, user_filename, app)
        database = create_database(user_filename, app)
        results_file = connect_database(query, converted_csv,
                                        database, user_filename, app)

        msg = Message(subject="Your Forms-to-Data Results",
                      sender=app.config.get("MAIL_USERNAME"),
                      recipients=[user_email],
                      body="Your results are attached!")
        with app.open_resource(results_file) as fp:
            msg.attach('results.csv', 'text/csv', fp.read())
        mail.send(msg)

        os.remove(xl)
        os.remove(query)
        os.remove(converted_csv)
        os.remove(database)
        os.remove(results_file)

        return render_template("sent.html")

    else:
        return render_template("convert.html")


@app.route("/help", methods=["GET"])
def register():
    """Info on files, various help"""
    return render_template("help.html")


def errorhandler(e):
    """Handle error"""
    if not isinstance(e, HTTPException):
        e = InternalServerError()
    return render_template("apology.html")


# Listen for errors
for code in default_exceptions:
    app.errorhandler(code)(errorhandler)
