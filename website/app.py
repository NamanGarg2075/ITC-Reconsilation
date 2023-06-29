from flask import Flask, render_template, request, send_file, make_response

# import pandas as pd
# import os
# import xlsxwriter

from asp2b import asp2b

from asp2a import asp2a

app = Flask(__name__)


@app.route("/")
def Home():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    try:
        action = request.form["action"]

        if action == "2B":
            return asp2b(app)
        if action == "2A":
            return asp2a(app)

    except Exception as e:
        return str(e)


if __name__ == "__main__":
    app.run(debug=True)
