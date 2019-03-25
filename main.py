import cs50
import re
import os
from flask import Flask, abort, redirect, render_template, request
from flask import send_file
from html import escape
from werkzeug.exceptions import default_exceptions, HTTPException
from werkzeug import secure_filename
from helpers import midlands, klinicare, rentmeester


# If `entrypoint` is not defined in app.yaml, App Engine will look for an app
# called `app` in `main.py`.
app = Flask(__name__)


@app.after_request
def after_request(response):
    """Disable caching"""
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Expires"] = 0
    response.headers["Pragma"] = "no-cache"
    return response


@app.route("/")
def index():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("index.html")


@app.route("/projects.html")
def projects():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("projects.html")


@app.route("/automate.html")
def automate():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("automate.html")


@app.route("/technical.html")
def technical():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("technical.html")


@app.route("/product.html")
def product():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("product.html")


@app.route("/success.html")
def success():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("success.html")


@app.route("/surveyform.html")
def surveyform():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("surveyform.html")


@app.route("/tribute.html")
def tribute():
    """Handle requests for / via GET (and POST)"""
    fileName = ""
    return render_template("tribute.html")


@app.route("/processed", methods=["GET", "POST"])
def processed():
    """Handle requests for /process via GET and POST"""
    # Upload file
    if not request.files["file"]:
        return render_template("inputFile.html")
    try:
        file = request.files["file"]
        if not file:
            return render_template("inputFile.html")
        file.save(secure_filename(file.filename))
    except Exception:
        return render_template("inputFile.html")
    # Default definition of fileName
    fileName = "Error.txt"
    # Process files
    if not request.form.get("algorithm"):
        return render_template("inputMethod.html")
    elif request.form.get("algorithm") == "midlands":
        fileName = midlands(file.filename)
        if fileName == "Error.txt":
            return render_template("failure.html")
        else:
            # Returns a webpage after successfully processing the file
            return render_template("successful.html", fileName=fileName)
    elif request.form.get("algorithm") == "klinicare":
        fileName = klinicare(file.filename)
        if fileName == "Error.txt":
            return render_template("failure.html")
        else:
            # Returns a webpage after successfully processing the file
            return render_template("successful.html", fileName=fileName)
    elif request.form.get("algorithm") == "rentmeester":
        fileName = rentmeester(file.filename)
        if fileName == "Error.txt":
            return render_template("failure.html")
        else:
            # Returns a webpage after successfully processing the file
            return render_template("successful.html", fileName=fileName)
    else:
        return render_template("inputMethod.html")


@app.route("/<fileName>")
def download(fileName):
    return send_file(fileName, as_attachment=True, attachment_filename=fileName)


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=8080, debug=True)
