import os

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session
)
from flask_session import Session
from tempfile import mkdtemp
from werkzeug.exceptions import (
    default_exceptions,
    HTTPException,
    InternalServerError
)

from helpers import (
    excel_convert,
    create_database,
    connect_database
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
        """Do stuff"""
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
