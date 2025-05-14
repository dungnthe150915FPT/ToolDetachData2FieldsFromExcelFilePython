# app.py
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

UPLOAD_FOLDER = 'static/uploads'
DOWNLOAD_FOLDER = 'static/downloads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

EMAIL_MAPPING = {
    "01_Công ty EMS": "quoctb@vnpost.vn",
		"02_Công ty VCKV": "quoctb@vnpost.vn",
		"03_Công ty PHBC": "quoctb@vnpost.vn",
		"04_Văn phòng Tổng công ty": "quoctb@vnpost.vn",
		"06_Công ty Dịch vụ số": "viettt@vnpost.vn",
		"08_Bưu điện Trung Ương": "viettt@vnpost.vn",
		"10_Bưu điện TP Hà Nội": "it.bdhn@vnpost.vn",
		"11_Bưu điện Trung tâm Hoàn Kiếm": "it.hoankiem@vnpost.vn",
		"12_Bưu điện Trung tâm Hà Đông": "it.hadong@vnpost.vn",
		"13_Bưu điện Trung tâm Cầu Giấy": "it.caugiay@vnpost.vn",
		"14_Bưu điện Trung tâm Từ Liêm": "it.tuliem@vnpost.vn",
		"15_Bưu điện Trung tâm Long Biên": "it.longbien@vnpost.vn",
		"16_Bưu điện Tỉnh Hưng Yên": "it.bdhy@vnpost.vn",
		"18_Bưu điện TP Hải Phòng": "it.bdhp@vnpost.vn",
		"1H_1H Test": "",
		"20_Bưu điện Tỉnh Quảng Ninh": "it.bdqnh@vnpost.vn",
		"21_Bưu điện Trung tâm Đông Anh": "it.donganh@vnpost.vn",
		"22_Bưu điện Tỉnh Bắc Ninh": "it.bdbn@vnpost.vn",
		"23_Bưu điện Tỉnh Bắc Giang": "it.bdbg@vnpost.vn",
		"24_Bưu điện Tỉnh Lạng Sơn": "it.bdls@vnpost.vn",
		"25_Bưu điện Tỉnh Thái Nguyên": "it.bdtn@vnpost.vn",
		"26_Bưu điện Tỉnh Bắc Kạn": "it.bdbk@vnpost.vn",
		"27_Bưu điện Tỉnh Cao Bằng": "it.bdcb@vnpost.vn",
		"28_Bưu điện Tỉnh Vĩnh Phúc": "it.bdvp@vnpost.vn",
		"29_Bưu điện Tỉnh Phú Thọ": "it.bdpt@vnpost.vn",
		"30_Bưu điện Tỉnh Tuyên Quang": "it.bdtq@vnpost.vn",
		"31_Bưu điện Tỉnh Hà Giang": "it.bdhg@vnpost.vn",
		"32_Bưu điện Tỉnh Yên Bái": "it.bdyb@vnpost.vn",
		"33_Bưu điện Tỉnh Lào Cai": "it.bdlci@vnpost.vn",
		"34_Bưu điện Trung tâm Thanh Trì": "it.thanhtri@vnpost.vn",
		"35_Bưu điện Tỉnh Hoà Bình": "it.bdhb@vnpost.vn",
		"36_Bưu điện Tỉnh Sơn La": "it.bdsl@vnpost.vn",
		"37_Bưu điện Trung tâm Chương Mỹ": "it.chuongmy@vnpost.vn",
		"38_Bưu điện Tỉnh Điện Biên": "it.bddb@vnpost.vn",
		"39_Bưu điện Tỉnh Lai Châu": "it.bdlc@vnpost.vn",
		"40_Bưu điện Tỉnh Hà Nam": "it.bdhnm@vnpost.vn",
		"41_Bưu điện Tỉnh Thái Bình": "it.bdtb@vnpost.vn",
		"42_Bưu điện Tỉnh Nam Định": "it.bdnd@vnpost.vn",
		"43_Bưu điện Tỉnh Ninh Bình": "it.bdnb@vnpost.vn",
		"44_Bưu điện Tỉnh Thanh Hoá": "it.bdth@vnpost.vn",
		"46_Bưu điện Tỉnh Nghệ An": "it.bdna@vnpost.vn",
		"48_Bưu điện Tỉnh Hà Tĩnh": "it.bdht@vnpost.vn",
		"49_Bưu điện Trung tâm Sơn Tây": "it.sontay@vnpost.vn",
		"51_Bưu điện Tỉnh Quảng Bình": "it.bdqb@vnpost.vn",
		"52_Bưu điện Tỉnh Quảng Trị": "it.bdqt@vnpost.vn",
		"53_Bưu điện Tỉnh Thừa Thiên Huế": "it.bdtth@vnpost.vn",
		"55_Bưu điện TP Đà Nẵng": "it.bddn@vnpost.vn",
		"56_Bưu điện Tỉnh Quảng Nam": "it.bdqn@vnpost.vn",
		"57_Bưu điện Tỉnh Quảng Ngãi": "it.bdqni@vnpost.vn",
		"58_Bưu điện Tỉnh Kon Tum": "it.bdkt@vnpost.vn",
		"59_Bưu điện Tỉnh Bình Định": "it.bdbdh@vnpost.vn",
		"60_Bưu điện Tỉnh Gia Lai": "it.bdgl@vnpost.vn",
		"62_Bưu điện Tỉnh Phú Yên": "it.bdpy@vnpost.vn",
		"63_Bưu điện Tỉnh Đắk Lăk": "it.bddl@vnpost.vn",
		"64_Bưu điện Tỉnh Đắk Nông": "it.bddng@vnpost.vn",
		"65_Bưu điện Tỉnh Khánh Hoà": "it.bdkh@vnpost.vn",
		"66_Bưu điện Tỉnh Ninh Thuận": "it.bdnt@vnpost.vn",
		"67_Bưu điện Tỉnh Lâm Đồng": "it.bdld@vnpost.vn",
		"70_Bưu điện TP Hồ Chí Minh": "it.bdhcm@vnpost.vn",
		"71_Bưu điện Trung tâm Sài Gòn": "it.saigon@vnpost.vn",
		"72_Bưu điện Trung tâm Phú Thọ": "it.phutho@vnpost.vn",
		"73_Bưu điện Trung tâm Chợ Lớn": "it.cholon@vnpost.vn",
		"74_Bưu điện Trung tâm Nam Sài Gòn": "it.namsaigon@vnpost.vn",
		"75_Bưu điện Trung tâm Gia Định": "it.giadinh@vnpost.vn",
		"76_Bưu điện Trung tâm Bình Chánh": "it.binhchanh@vnpost.vn",
		"77_Bưu điện Trung tâm Củ Chi": "it.cuchi@vnpost.vn",
		"78_Bưu điện Thành phố Thủ Đức": "it.thuduc@vnpost.vn",
		"80_Bưu điện Tỉnh Bình Thuận": "it.bdbt@vnpost.vn",
		"81_Bưu điện Tỉnh Đồng Nai": "it.bddni@vnpost.vn",
		"82_Bưu điện Tỉnh Bình Dương": "it.bdbd@vnpost.vn",
		"84_Bưu điện Tỉnh Tây Ninh": "it.bdtn@vnpost.vn",
		"85_Bưu điện Tỉnh Long An": "it.bdla@vnpost.vn",
		"86_Bưu điện Tỉnh Tiền Giang": "it.bdtg@vnpost.vn",
		"87_Bưu điện Tỉnh Đồng Tháp": "it.bddt@vnpost.vn",
		"88_Bưu điện Tỉnh An Giang": "it.bdag@vnpost.vn",
		"89_Bưu điện Tỉnh Vĩnh Long": "it.bdvl@vnpost.vn",
		"90_Bưu điện TP Cần Thơ": "it.bdct@vnpost.vn",
		"91_Bưu điện Tỉnh Hậu Giang": "it.bdhgg@vnpost.vn",
		"92_Bưu điện Tỉnh Kiên Giang": "it.bdkg@vnpost.vn",
		"93_Bưu điện Tỉnh Bến Tre": "it.bdbt@vnpost.vn",
		"94_Bưu điện Tỉnh Trà Vinh": "it.bdtv@vnpost.vn",
		"95_Bưu điện Tỉnh Sóc Trăng": "it.bdst@vnpost.vn",
		"96_Bưu điện Tỉnh Bạc Liêu": "it.bdbl@vnpost.vn",
		"97_Bưu điện Tỉnh Cà Mau": "it.bdcm@vnpost.vn",
		"HN_Hà Nội Test": "viettt@vnpost.vn"
}

def sanitize_filename(name):
    name = str(name).replace('_', ' ').strip()
    return re.sub(r'[\\/*?:"<>|]', '', name)

@app.route('/', methods=['GET', 'POST'])
def index():
    status = ""
    download_links = []

    if request.method == 'POST':
        try:
            uploaded_file = request.files.get('file')
            sender_email = request.form.get('sender_email')
            sender_password = request.form.get('sender_password')
            subject_template = request.form.get('email_subject', 'DANH SÁCH TÀI KHOẢN ĐANG ĐỂ PASSWORD MẶC ĐỊNH {unit}')
            body_template = request.form.get('email_body', 'Kính gửi đơn vị {unit},\n\nEm gửi danh sách các tài khoản đang để password mặc định của đơn vị ngày 14/05/2025. Nhờ a/c đôn đốc xử lý thay đổi mật khẩu giúp em ạ.\n\nTrân trọng!')

            current_mapping = {}
            for org in EMAIL_MAPPING.keys():
                email_key = f"email_{org}"
                current_mapping[org] = request.form.get(email_key, "").strip()

            if not uploaded_file or uploaded_file.filename == '':
                raise ValueError("❌ Không có file nào được tải lên!")

            filename = secure_filename(uploaded_file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            uploaded_file.save(file_path)

            try:
                df = pd.read_excel(file_path, dtype=str)
            except Exception as e:
                raise ValueError(f"❌ Lỗi khi đọc file Excel: {e}")

            data_groups = defaultdict(list)

            for _, row in df.iterrows():
                org = str(row.get('ORG_CODE_NAME_BDT', '')).strip()
                user = str(row.get('USERNAME', '')).strip()
                if org and user:
                    data_groups[org].append([user, org])

            if not data_groups:
                raise ValueError("❌ File không chứa dữ liệu hợp lệ!")

            zip_filename = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')
            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for org, users in data_groups.items():
                    safe_org = sanitize_filename(org)
                    out_filename = f"{safe_org}.xlsx"
                    out_path = os.path.join(app.config['DOWNLOAD_FOLDER'], out_filename)

                    wb = Workbook()
                    ws = wb.active
                    ws.append(["USERNAME", "ORG_CODE_NAME_BDT"])
                    for user in users:
                        ws.append(user)
                    wb.save(out_path)

                    zipf.write(out_path, arcname=out_filename)
                    download_links.append(f"downloads/{out_filename}")

                    recipient_email = current_mapping.get(org)
                    if recipient_email:
                        subject = subject_template.replace("{unit}", org)
                        body = body_template.replace("{unit}", org)
                        send_status = send_email(sender_email, sender_password, recipient_email, subject, body, out_path)
                        print(send_status)

            status = "✅ Tách file và gửi mail thành công"

        except ValueError as e:
            status = str(e)
        except Exception as e:
            status = f"❌ Lỗi hệ thống: {e}"

        return render_template('index.html', status=status, download_links=download_links, email_mapping=current_mapping)

    return render_template('index.html', email_mapping=EMAIL_MAPPING)

def send_email(sender_email, sender_password, recipient_email, subject, body, attachment_path):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg.set_content(body)

        with open(attachment_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=os.path.basename(attachment_path))

        smtp_server = "mail.vnpost.vn"
        smtp_port = 587

        with smtplib.SMTP(smtp_server, smtp_port, timeout=10) as smtp:
            smtp.starttls()
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)

        os.remove(attachment_path)
        return f"✅ Email đã gửi đến {recipient_email}."

    except Exception as e:
        return f"❌ Gửi mail lỗi: {e}"

@app.route('/download-all')
def download_all():
    zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], 'all_groups.zip')
    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
