from flask import Flask, render_template, request, send_file
import os
import zipfile
import re
import smtplib
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from werkzeug.utils import secure_filename
from email.message import EmailMessage
from dotenv import load_dotenv

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['DOWNLOAD_FOLDER'] = 'static/downloads'

load_dotenv()  # Load biến môi trường email

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)


def sanitize_filename(name):
    """Loại bỏ ký tự không hợp lệ, giữ lại tiếng Việt và khoảng trắng"""
    name = name.replace('_', ' ')  # Nếu org chứa dấu gạch dưới, đổi thành dấu cách
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()


@app.route('/', methods=['GET', 'POST'])
def index():
    status = ""
    download_links = []

    if request.method == 'POST':
        uploaded_file = request.files.get('file')
        email = request.form.get('email')

        if not uploaded_file or uploaded_file.filename == '':
            status = "Không có file được tải lên."
            return render_template('index.html', status=status)

        filename = secure_filename(uploaded_file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        uploaded_file.save(file_path)

        wb = load_workbook(file_path)
        ws = wb.active
        data = defaultdict(list)

        for row in ws.iter_rows(min_row=2, values_only=True):
            username, org = row[:2]
            if username and org:
                data[org].append((username, org))

        zip_filename = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for org, users in data.items():
                safe_org = sanitize_filename(org)
                out_filename = f"{safe_org}.xlsx"
                out_path = os.path.join(app.config['DOWNLOAD_FOLDER'], out_filename)

                out_wb = Workbook()
                out_ws = out_wb.active
                out_ws.append(["USERNAME", "ORG_CODE_NAME_BDT"])
                for user in users:
                    out_ws.append(user)
                out_wb.save(out_path)

                zipf.write(out_path, arcname=out_filename)
                download_links.append(f"downloads/{out_filename}")

        # Gửi email nếu có
        if email:
            try:
                msg = EmailMessage()
                msg['Subject'] = "Tách file hoàn tất"
                msg['From'] = os.getenv("EMAIL_USER")
                msg['To'] = email
                msg.set_content("Đã xử lý và đính kèm file.")

                with open(zip_filename, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='zip', filename='all_groups.zip')

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
                    smtp.send_message(msg)

                status = f"✅ Đã gửi file tới {email}"
            except Exception as e:
                status = f"❌ Lỗi gửi email: {e}"

        return render_template('index.html', status=status, download_links=download_links)

    return render_template('index.html')


@app.route('/download-all')
def download_all():
    zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')
    return send_file(zip_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
