from flask import Flask, render_template, request, send_file
import os
import zipfile
import re
import smtplib
from collections import defaultdict
from openpyxl import Workbook
from werkzeug.utils import secure_filename
from email.message import EmailMessage
import pandas as pd

app = Flask(__name__)

# Cấu hình thư mục lưu trữ file
UPLOAD_FOLDER = 'static/uploads'
DOWNLOAD_FOLDER = 'static/downloads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# Danh sách email cố định cho từng tổ chức
EMAIL_MAPPING = {
    "22_Bưu điện Tỉnh Bắc Ninh": "dungnguyentuan2001@gmail.com",
    "23_Bưu điện Tỉnh Bắc Giang": "dungdev224@gmail.com",
    "24_Bưu điện Tỉnh Hà Giang": "vipdungntd224@gmail.com",
}


def sanitize_filename(name):
    """Loại bỏ ký tự không hợp lệ khỏi tên file"""
    name = str(name).replace('_', ' ').strip()
    return re.sub(r'[\\/*?:"<>|]', '', name)


@app.route('/', methods=['GET', 'POST'])
# def index():
#     status = ""
#     download_links = []

#     if request.method == 'POST':
#         uploaded_file = request.files.get('file')
#         sender_email = request.form.get('sender_email')
#         sender_password = request.form.get('sender_password')
#         subject_template = request.form.get('email_subject', 'Báo cáo tổ chức: {unit}')
#         body_template = request.form.get('email_body', 'Xin chào,\n\nĐính kèm là báo cáo của {unit}.')

#         if not uploaded_file or uploaded_file.filename == '':
#             status = "❌ Không có file nào được tải lên!"
#             return render_template('index.html', status=status)

#         # Lưu file Excel được tải lên
#         filename = secure_filename(uploaded_file.filename)
#         file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
#         uploaded_file.save(file_path)

#         try:
#             df = pd.read_excel(file_path, dtype=str)
#         except Exception as e:
#             status = f"❌ Lỗi khi đọc file Excel: {e}"
#             return render_template('index.html', status=status)

#         data_groups = defaultdict(list)

#         # Chia dữ liệu thành nhóm dựa trên cột ORG_CODE_NAME_BDT
#         for _, row in df.iterrows():
#             org = str(row.get('ORG_CODE_NAME_BDT', '')).strip()
#             user = str(row.get('USERNAME', '')).strip()
#             if org and user:
#                 data_groups[org].append([user, org])

#         # Tạo file ZIP chứa tất cả nhóm
#         zip_filename = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')

#         with zipfile.ZipFile(zip_filename, 'w') as zipf:
#             for org, users in data_groups.items():
#                 safe_org = sanitize_filename(org)
#                 out_filename = f"{safe_org}.xlsx"
#                 out_path = os.path.join(app.config['DOWNLOAD_FOLDER'], out_filename)

#                 # Tạo file Excel cho từng nhóm
#                 wb = Workbook()
#                 ws = wb.active
#                 ws.append(["USERNAME", "ORG_CODE_NAME_BDT"])
#                 for user in users:
#                     ws.append(user)
#                 wb.save(out_path)

#                 # Thêm vào file ZIP
#                 zipf.write(out_path, arcname=out_filename)
#                 download_links.append(f"downloads/{out_filename}")

#                 # Gửi email nếu có địa chỉ tương ứng
#                 recipient_email = EMAIL_MAPPING.get(org)
#                 if recipient_email:
#                     subject = subject_template.replace("{unit}", org)
#                     body = body_template.replace("{unit}", org)
#                     send_status = send_email(sender_email, sender_password, recipient_email, subject, body, out_path)
#                     print(send_status)  # Ghi log trạng thái gửi email

        

#     return render_template('index.html')

@app.route('/', methods=['GET', 'POST'])
def index():
    status = ""
    download_links = []

    if request.method == 'POST':
        try:
            uploaded_file = request.files.get('file')
            sender_email = request.form.get('sender_email')
            sender_password = request.form.get('sender_password')
            subject_template = request.form.get('email_subject', 'Báo cáo tổ chức: {unit}')
            body_template = request.form.get('email_body', 'Xin chào,\n\nĐính kèm là báo cáo của {unit}.')

            if not uploaded_file or uploaded_file.filename == '':
                raise ValueError("❌ Không có file nào được tải lên!")

            # Lưu file Excel được tải lên
            filename = secure_filename(uploaded_file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            uploaded_file.save(file_path)

            try:
                df = pd.read_excel(file_path, dtype=str)
            except Exception as e:
                raise ValueError(f"❌ Lỗi khi đọc file Excel: {e}")

            data_groups = defaultdict(list)

            # Chia dữ liệu thành nhóm dựa trên cột ORG_CODE_NAME_BDT
            for _, row in df.iterrows():
                org = str(row.get('ORG_CODE_NAME_BDT', '')).strip()
                user = str(row.get('USERNAME', '')).strip()
                if org and user:
                    data_groups[org].append([user, org])

            if not data_groups:
                raise ValueError("❌ File không chứa dữ liệu hợp lệ!")

            # Tạo file ZIP chứa tất cả nhóm
            zip_filename = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')

            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for org, users in data_groups.items():
                    safe_org = sanitize_filename(org)
                    out_filename = f"{safe_org}.xlsx"
                    out_path = os.path.join(app.config['DOWNLOAD_FOLDER'], out_filename)

                    # Tạo file Excel cho từng nhóm
                    wb = Workbook()
                    ws = wb.active
                    ws.append(["USERNAME", "ORG_CODE_NAME_BDT"])
                    for user in users:
                        ws.append(user)
                    wb.save(out_path)

                    # Thêm vào file ZIP
                    zipf.write(out_path, arcname=out_filename)
                    download_links.append(f"downloads/{out_filename}")

                    # Gửi email nếu có địa chỉ tương ứng
                    recipient_email = EMAIL_MAPPING.get(org)
                    if recipient_email:
                        subject = subject_template.replace("{unit}", org)
                        body = body_template.replace("{unit}", org)
                        send_status = send_email(sender_email, sender_password, recipient_email, subject, body, out_path)
                        print(send_status)  # Ghi log trạng thái gửi email

            status = "✅ Xử lý xong! Đã gửi email nếu có địa chỉ được chỉ định."

        except ValueError as e:
            status = str(e)
        except Exception as e:
            status = f"❌ Lỗi hệ thống: {e}"

        return render_template('index.html', status=status, download_links=download_links)

    return render_template('index.html')

def send_email(sender_email, sender_password, recipient_email, subject, body, attachment_path):
    """Gửi email với file đính kèm và kiểm tra lỗi"""
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg.set_content(body)

        with open(attachment_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=os.path.basename(attachment_path))

        smtp_server = "mail.vnpost.vn"
        smtp_port = 587

        with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as smtp:
            smtp.starttls()
            try:
                smtp.login(sender_email, sender_password)
            except smtplib.SMTPAuthenticationError:
                return "❌ Lỗi xác thực! Kiểm tra email/mật khẩu."
            except smtplib.SMTPConnectError:
                return "❌ Không thể kết nối đến máy chủ SMTP!"
            except smtplib.SMTPException as e:
                return f"❌ Lỗi SMTP: {e}"

            try:
                smtp.send_message(msg)
            except smtplib.SMTPRecipientsRefused:
                return f"❌ Email bị từ chối! Kiểm tra lại ({recipient_email})."
            except smtplib.SMTPException as e:
                return f"❌ Lỗi khi gửi email: {e}"

        os.remove(attachment_path)  # Xóa file sau khi gửi thành công
        return f"✅ Email đã gửi đến {recipient_email}."

    except Exception as e:
        return f"❌ Lỗi hệ thống khi gửi email: {e}"


@app.route('/download-all')
def download_all():
    """Cho phép tải xuống tất cả các file đã xử lý dưới dạng ZIP"""
    zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')
    return send_file(zip_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
