import os
import sqlite3
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_file, g
)
from werkzeug.security import generate_password_hash, check_password_hash

from utils import (
    ensure_dir, random_filename, convert_office_to_pdf,
    allowed_file, merge_pdfs, split_pdf
)

# ------------------------------------------
# CONFIG
# ------------------------------------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
CONVERTED_DIR = os.path.join(BASE_DIR, "converted")
DB_PATH = os.path.join(BASE_DIR, "app.db")

SECRET_KEY = "super-secret-key"

for folder in (UPLOAD_DIR, CONVERTED_DIR):
    ensure_dir(folder)

app = Flask(__name__)
app.config["SECRET_KEY"] = SECRET_KEY
app.config["UPLOAD_FOLDER"] = UPLOAD_DIR


# ------------------------------------------
# DATABASE CONNECTION
# ------------------------------------------
def get_db():
    db = getattr(g, "_database", None)
    if db is None:
        db = g._database = sqlite3.connect(DB_PATH)
        db.row_factory = sqlite3.Row
    return db


@app.teardown_appcontext
def close_db(exception):
    db = getattr(g, "_database", None)
    if db:
        db.close()


# ------------------------------------------
# LOGIN REQUIRED DECORATOR
# ------------------------------------------
def login_required(fn):
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return fn(*args, **kwargs)

    wrapper.__name__ = fn.__name__
    return wrapper


# ------------------------------------------
# AUTH ROUTES
# ------------------------------------------
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        uname = request.form["username"]
        pwd = request.form["password"]

        if not uname or not pwd:
            flash("Fill all fields", "danger")
            return redirect(url_for("register"))

        db = get_db()
        try:
            db.execute(
                "INSERT INTO users (username, password_hash) VALUES (?, ?)",
                (uname, generate_password_hash(pwd))
            )
            db.commit()
            flash("Registered successfully", "success")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("Username already exists", "danger")

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        uname = request.form["username"]
        pwd = request.form["password"]

        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE username = ?",
            (uname,)
        ).fetchone()

        if user and check_password_hash(user["password_hash"], pwd):
            session["user_id"] = user["id"]
            return redirect(url_for("dashboard"))

        flash("Invalid username or password", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ------------------------------------------
# DASHBOARD
# ------------------------------------------
@app.route("/")
@login_required
def dashboard():
    return render_template("dashboard.html")


# ------------------------------------------
# WORD → PDF
# ------------------------------------------
@app.route("/convert/word", methods=["GET", "POST"])
@login_required
def convert_word():
    if request.method == "POST":
        file = request.files.get("file")

        if file and file.filename.lower().endswith(".docx"):
            filename = random_filename(file.filename)
            save_path = os.path.join(UPLOAD_DIR, filename)
            file.save(save_path)

            output_file = convert_office_to_pdf(save_path, CONVERTED_DIR)
            return send_file(output_file, as_attachment=True)

        flash("Upload a valid .docx file!", "danger")

    return render_template("word_to_pdf.html")


# ------------------------------------------
# PPT → PDF
# ------------------------------------------
@app.route("/convert/ppt", methods=["GET", "POST"])
@login_required
def convert_ppt():
    if request.method == "POST":
        file = request.files.get("file")

        if file and file.filename.lower().endswith((".ppt", ".pptx")):
            filename = random_filename(file.filename)
            save_path = os.path.join(UPLOAD_DIR, filename)
            file.save(save_path)

            output_file = convert_office_to_pdf(save_path, CONVERTED_DIR)
            return send_file(output_file, as_attachment=True)

        flash("Upload a valid PPT or PPTX file!", "danger")

    return render_template("ppt_to_pdf.html")


# ------------------------------------------
# IMAGE → PDF
# ------------------------------------------
@app.route("/convert/image", methods=["GET", "POST"])
@login_required
def convert_image():
    if request.method == "POST":
        file = request.files.get("file")

        if file and file.filename.lower().endswith((".jpg", ".jpeg", ".png")):
            filename = random_filename(file.filename)
            save_path = os.path.join(UPLOAD_DIR, filename)
            file.save(save_path)

            output_path = os.path.join(CONVERTED_DIR, random_filename("image.pdf"))

            from utils import image_to_pdf
            image_to_pdf(save_path, output_path)

            return send_file(output_path, as_attachment=True)

        flash("Upload a JPG, JPEG, or PNG image!", "danger")

    return render_template("image_to_pdf.html")


# ------------------------------------------
# EXCEL → PDF
# ------------------------------------------
@app.route("/convert/excel", methods=["GET", "POST"])
@login_required
def convert_excel():
    if request.method == "POST":
        file = request.files.get("file")

        if file and file.filename.lower().endswith((".xls", ".xlsx")):
            filename = random_filename(file.filename)
            save_path = os.path.join(UPLOAD_DIR, filename)
            file.save(save_path)

            output_file = convert_office_to_pdf(save_path, CONVERTED_DIR)
            return send_file(output_file, as_attachment=True)

        flash("Upload a valid XLS or XLSX file!", "danger")

    return render_template("excel_to_pdf.html")


# ------------------------------------------
# MERGE PDFs
# ------------------------------------------
@app.route("/merge", methods=["GET", "POST"])
@login_required
def merge():
    if request.method == "POST":
        files = request.files.getlist("files")

        if len(files) < 2:
            flash("Upload at least 2 PDFs", "danger")
            return redirect(url_for("merge"))

        file_paths = []
        for f in files:
            fname = random_filename(f.filename)
            path = os.path.join(UPLOAD_DIR, fname)
            f.save(path)
            file_paths.append(path)

        output_path = os.path.join(CONVERTED_DIR, random_filename("merged.pdf"))
        merge_pdfs(file_paths, output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("merge.html")


# ------------------------------------------
# SPLIT PDF
# ------------------------------------------
@app.route("/split", methods=["GET", "POST"])
@login_required
def split():
    if request.method == "POST":
        pdf = request.files.get("file")
        pages = request.form.get("pages")

        if not pdf or not pages:
            flash("Upload PDF and enter pages!", "danger")
            return redirect(url_for("split"))

        filename = random_filename(pdf.filename)
        path = os.path.join(UPLOAD_DIR, filename)
        pdf.save(path)

        page_list = []
        for part in pages.split(","):
            part = part.strip()
            if "-" in part:
                start, end = map(int, part.split("-"))
                page_list.extend(range(start, end + 1))
            else:
                page_list.append(int(part))

        output_path = os.path.join(CONVERTED_DIR, random_filename("split.pdf"))
        split_pdf(path, page_list, output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("split.html")


# ------------------------------------------
# COMPRESS (REMOVED)
# ------------------------------------------
# You requested to remove the compress function.


# ------------------------------------------
# UI PAGES
# ------------------------------------------
@app.route("/unlock")
@login_required
def unlock():
    return render_template("unlock.html")


@app.route("/ocr")
@login_required
def ocr():
    return render_template("ocr.html")


# ------------------------------------------
# RUN APP
# ------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)
