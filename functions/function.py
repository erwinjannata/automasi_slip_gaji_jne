import xlwings as xl
import re
import os
from dotenv import load_dotenv
from pathlib import Path
from tkinter.messagebox import showinfo
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_email(file_data):
    app = xl.App(visible=False)
    main_book = xl.Book(file_data)
    log_sheet = main_book.sheets['Log Email']
    data_sheet = main_book.sheets['Rekap JNE']
    template_sheet = main_book.sheets['SLIP TEMPLATE']
    periode = template_sheet['N6'].value
    max_row = int(re.findall(
        r'\d+', (data_sheet.range("B8").end("down").address))[0])

    working_directory = Path.cwd()

    load_dotenv()
    sender = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    recipent_name = data_sheet[f'B8:B{max_row}'].value
    recipent_email = data_sheet[f'AP8:AP{max_row}'].value
    subject = f'Slip Gaji Periode {periode}'
    body = 'Do not reply this email'

    for index, name in enumerate(recipent_name):
        if Path(f'{Path.cwd()}/{periode}/{name}.pdf').exists():
            with open(rf'{working_directory}\{periode}\{name}.pdf', 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition', f"attachment; filename={name}.pdf"
            )
        else:
            log_sheet[f'B{index+2}'].value = 'File Tidak Ditemukan'
            continue

        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = f'JNE AMI <{sender}>'
        message['To'] = recipent_email[index]
        html_part = MIMEText(body)
        message.attach(html_part)
        message.attach(part)

        if recipent_email[index] == None:
            log_sheet[f'B{index+2}'].value = 'Email Kosong'
            continue
        else:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                if not log_sheet[f'B{index+2}'].value == 'Terkirim':
                    try:
                        server.login(sender, sender_password)
                        server.sendmail(
                            sender, recipent_email[index], message.as_string())
                        log_sheet[f'B{index+2}'].value = 'Terkirim'
                    except Exception as e:
                        log_sheet[f'B{index+2}'].value = 'Gagal'
                    if index == (len(recipent_name) - 1):
                        server.close()

    main_book.save()
    main_book.close()
    app.quit()
    showinfo(title="Message",
             message="Pengiriman selesai, silahkan cek lembar Log Email")


def generate_slip(file_data):
    app = xl.App(visible=False)

    try:
        global main_book, data_sheet
        main_book = xl.Book(file_data)

        data_sheet = main_book.sheets['Rekap JNE']
        detail_sheet = main_book.sheets['DETAIL ']
        template_sheet = main_book.sheets['SLIP TEMPLATE']

        main_book.sheets.add(name='Log Email', after='Rekap JNE')
        log_sheet = main_book.sheets['Log Email']
        log_sheet['A1'].value = 'Nama'
        log_sheet['B1'].value = 'Status'

        periode = template_sheet['N6'].value
        working_directory = Path.cwd()/periode

        if Path.exists(working_directory) == False:
            working_directory.mkdir(parents=True, exist_ok=True)

        max_row = int(re.findall(
            r'\d+', (data_sheet.range("B8").end("down").address))[0])

        detail_max_pusat = int(re.findall(
            r'\d+', (detail_sheet.range('B4').end("down").address))[0])

        for i in range(8, max_row + 1):
            nama = data_sheet[f'B{i}'].value
            jabatan = data_sheet[f'C{i}'].value
            gaji_pokok = data_sheet[f'E{i}'].value or 0

            # Tunjangan
            uang_makan = data_sheet[f'K{i}'].value or 0
            tunjangan_jabatan = data_sheet[f'F{i}'].value or 0
            tunjangan_pendidikan = data_sheet[f'H{i}'].value or 0
            tunjangan_masa_kerja = data_sheet[f'G{i}'].value or 0
            tunjangan_beras = data_sheet[f'O{i}'].value or 0
            pulsa = data_sheet[f'I{i}'].value or 0
            piket = data_sheet[f'L{i}'].value or 0
            lembur = data_sheet[f'M{i}'].value or 0
            bonus = data_sheet[f'N{i}'].value or 0
            sewa_motor = data_sheet[f'P{i}'].value or 0
            claim_lalu = data_sheet[f'J{i}'].value or 0

            # BPJS
            if i-5 < detail_max_pusat:
                index = i - 4
            else:
                index = i + 5
            tunjangan_jpk = detail_sheet[f'X{index}'].value or 0
            tunjangan_jkm = detail_sheet[f'Y{index}'].value or 0
            tunjangan_jkk = detail_sheet[f'Z{index}'].value or 0
            tunjangan_jht = detail_sheet[f'AA{index}'].value or 0
            tunjangan_jpn = detail_sheet[f'AB{index}'].value or 0
            bpjs = (tunjangan_jpk + tunjangan_jkm +
                    tunjangan_jkk + tunjangan_jht + tunjangan_jpn)
            total_tunjangan = (gaji_pokok + uang_makan + tunjangan_jabatan + tunjangan_pendidikan +
                               tunjangan_masa_kerja + tunjangan_beras + pulsa + piket + lembur + bonus)

            # Potongan
            potongan_jpk = data_sheet[f'AG{i}'].value or 0
            potongan_jht = data_sheet[f'AF{i}'].value or 0
            kasbon = data_sheet[f'AI{i}'].value or 0

            n_alpa = detail_sheet[f'AC{index}'].value or 0
            n_cuti = detail_sheet[f'AD{index}'].value or 0
            n_sakit = detail_sheet[f'AE{index}'].value or 0
            n_set_hari = detail_sheet[f'AF{index}'].value or 0
            n_telat = detail_sheet[f'AG{index}'].value or 0
            n_cuti_habis = detail_sheet[f'AH{index}'].value or 0

            potongan_alpa = data_sheet[f'S{i}'].value or 0
            potongan_cuti_habis = data_sheet[f'X{i}'].value or 0
            potongan_cuti = data_sheet[f'T{i}'].value or 0
            potongan_sakit = data_sheet[f'U{i}'].value or 0
            potongan_set_hari = data_sheet[f'V{i}'].value or 0
            potongan_telat = data_sheet[f'W{i}'].value or 0
            potongan_claim_barang = data_sheet[f'AJ{i}'].value or 0
            potongan_claim = data_sheet[f'AK{i}'].value or 0
            potongan_sp = data_sheet[f'AL{i}'].value or 0
            potongan_lain = bpjs + claim_lalu + sewa_motor
            total_potongan = potongan_jpk + potongan_jht + kasbon + potongan_alpa + potongan_cuti_habis + \
                potongan_cuti + potongan_sakit + potongan_set_hari + potongan_telat + potongan_lain

            # Paste Tunjangan
            template_sheet['C6'].value = nama
            template_sheet['C7'].value = jabatan
            template_sheet['H10'].value = gaji_pokok
            template_sheet['H12'].value = uang_makan
            template_sheet['H13'].value = tunjangan_jabatan
            template_sheet['H14'].value = tunjangan_pendidikan
            template_sheet['H15'].value = tunjangan_masa_kerja
            template_sheet['H16'].value = tunjangan_beras
            template_sheet['H17'].value = pulsa
            template_sheet['H18'].value = tunjangan_jpk
            template_sheet['H19'].value = tunjangan_jkm
            template_sheet['H20'].value = tunjangan_jkk
            template_sheet['H21'].value = tunjangan_jht
            template_sheet['H22'].value = tunjangan_jpn
            template_sheet['H23'].value = piket
            template_sheet['H24'].value = lembur
            template_sheet['H25'].value = bonus
            template_sheet['H26'].value = sewa_motor
            template_sheet['H27'].value = claim_lalu
            template_sheet['H28'].value = total_tunjangan + potongan_lain

            # Paste Potongan
            template_sheet['O12'].value = potongan_jpk
            template_sheet['O13'].value = potongan_jht
            template_sheet['O14'].value = kasbon
            template_sheet['L15'].value = n_alpa
            template_sheet['O15'].value = potongan_alpa
            template_sheet['L16'].value = n_cuti_habis
            template_sheet['O16'].value = potongan_cuti_habis
            template_sheet['L17'].value = n_cuti
            template_sheet['O17'].value = potongan_cuti
            template_sheet['L18'].value = n_sakit
            template_sheet['O18'].value = potongan_sakit
            template_sheet['L19'].value = n_set_hari
            template_sheet['O19'].value = potongan_set_hari
            template_sheet['L20'].value = n_telat
            template_sheet['O20'].value = potongan_telat
            template_sheet['O21'].value = potongan_sp
            template_sheet['O22'].value = potongan_claim_barang
            template_sheet['O23'].value = potongan_claim
            template_sheet['O24'].value = potongan_lain
            template_sheet['O25'].value = total_potongan

            template_sheet['O31'].value = data_sheet[f'AN{i}'].value
            template_sheet['O32'].value = data_sheet[f'AO{i}'].value
            template_sheet['O36'].value = data_sheet[f'AN{i}'].value + \
                data_sheet[f'AO{i}'].value

            template_sheet['F36'].value = nama

            # Log Email
            log_sheet[f'A{i-6}'].value = nama
            log_sheet[f'B{i-6}'].value = ''

            template_sheet.to_pdf(
                path=rf'{working_directory}/{nama}.pdf', quality='standard')

        main_book.save()
        main_book.close()
        app.quit()
        showinfo(title="Message",
                 message=f"Pembuatan Slip Gaji periode {periode} selesai")
    except OSError:
        app.quit()
        showinfo(title="Message",
                 message="File excel sedang dibuka / digunakan oleh proses lain.")
    except Exception as e:
        main_book.close()
        app.quit()
        showinfo(title="Message",
                 message=f"{e}")
