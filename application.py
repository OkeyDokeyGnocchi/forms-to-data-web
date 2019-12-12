import os

from flask import (
    Flask,
    redirect,
    render_template,
    request,
    session
)
from flask_session import Session
from flask_mail import (
    Mail,
    Message
)
from tempfile import mkdtemp
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
app = Flask(__name__)

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
        user_input = request.form

        if "xlsxFile" not in user_input.files:
            return render_template("apology.html")
        elif "queryList" not in user_input.files:
            return render_template("apology.html")
        else:
            xlsxFile = request.files['xlsxFile']
            user_query = request.files["queryList"]
            user_email = request.form.get("userEmail")
            user_filename = generate_filename()

        converted_csv = excel_convert(xlsxFile, user_filename)
        database = create_database(user_filename)
        results_file = connect_database(user_query, converted_csv,
                                        database, user_filename)

        msg = Message(subject="Your Forms-to-Data Results",
                      sender=app.config.get("MAIL_USERNAME"),
                      recipients=[user_email],
                      body="Your results are attached!")
        msg.attach(results_file)
        mail.send(msg)

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
