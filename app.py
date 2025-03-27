from flask import Flask, render_template, request, send_file
import os
import zipfile
import re
import smtplib
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from werkzeug.utils import secure_filename
from email.message import EmailMessage
import pandas as pd  # Thêm pandas để xử lý Excel nhanh hơn

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['DOWNLOAD_FOLDER'] = 'static/downloads'

# Tạo thư mục nếu chưa có
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)


def sanitize_filename(name):
    """Loại bỏ ký tự không hợp lệ, giữ lại tiếng Việt và khoảng trắng"""
    name = str(name)  # Đảm bảo dữ liệu là chuỗi
    name = name.replace('_', ' ')  # Nếu có dấu gạch dưới, đổi thành dấu cách
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()


@app.route('/', methods=['GET', 'POST'])
def index():
    status = ""
    download_links = []

    if request.method == 'POST':
        uploaded_file = request.files.get('file')
        recipient_email = request.form.get('email')
        sender_email = request.form.get('sender_email')
        sender_password = request.form.get('sender_password')

        if not uploaded_file or uploaded_file.filename == '':
            status = "Không có file được tải lên."
            return render_template('index.html', status=status)

        filename = secure_filename(uploaded_file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        uploaded_file.save(file_path)

        # Đọc file Excel
        try:
            df = pd.read_excel(file_path, dtype=str)  # Đọc file, mọi dữ liệu là chuỗi
        except Exception as e:
            status = f"Lỗi khi đọc file Excel: {e}"
            return render_template('index.html', status=status)

        # Tạo dictionary để lưu dữ liệu theo nhóm tổ chức
        data = defaultdict(list)

        for _, row in df.iterrows():
            org = str(row.get('ORG_CODE_NAME_BDT', '')).strip()
            user = str(row.get('USERNAME', '')).strip()
            if org and user:  # Bỏ qua dòng trống
                data[org].append([user, org])

        # Đường dẫn file zip chứa tất cả file Excel đã tách
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

        # **Gửi email bằng tài khoản của người dùng**
        if recipient_email and sender_email and sender_password:
            try:
                msg = EmailMessage()
                msg['Subject'] = "Tách file hoàn tất"
                msg['From'] = sender_email
                msg['To'] = recipient_email
                msg.set_content("Đã xử lý và đính kèm file.")

                with open(zip_filename, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='zip', filename='all_groups.zip')

                # **Kết nối SMTP của công ty**
                smtp_server = "mail.vnpost.vn"
                smtp_port = 587

                with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                    smtp.starttls()  # Bật bảo mật TLS
                    smtp.login(sender_email, sender_password)
                    smtp.send_message(msg)

                status = f"✅ Đã gửi file tới {recipient_email}"
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
