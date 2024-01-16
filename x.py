import math
import pandas as pd
from tkinter import filedialog, messagebox, font, Text, Scrollbar, simpledialog
import tkinter as tk
from PIL import Image, ImageTk
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def send_email(gonderen_email, gonderen_sifre, alici_email, konu, mesaj, dosya_yolu):
    #E-posta ayarları
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    #E-posta gönderme işlemi
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(gonderen_email, gonderen_sifre)

        msg = MIMEMultipart()
        msg['From'] = gonderen_email
        msg['To'] = alici_email
        msg['Subject'] = konu

        body = mesaj
        msg.attach(MIMEText(body, 'plain'))

        #Eğer dosya varsa ekle
        if dosya_yolu:
            with open(dosya_yolu, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name="dosya.xlsx")

            part['Content-Disposition'] = f'attachment; filename={dosya_yolu}'
            msg.attach(part)

        server.sendmail(gonderen_email, alici_email, msg.as_string())

def Error_Detection(data1, data2):
    if data1 <= 0 and data2 >= 210:
        return "CURRENT WARNING"
    elif data1 > 0 and 80 <= data2 <= 100:
        return "FUSE WARNING"
    elif data2 == 0:
        return "VOLTAGE WARNING"

#TODO:Time koşul ekle
def Flow_Detector_With_Time(flow1,flow2,flow3,date,start_hour,end_hour):
    hour =float(date[11:16].replace(":", "."))
    start_hour=float(start_hour.replace(":","."))
    end_hour=float(end_hour.replace(":","."))

    if start_hour < hour < end_hour:
        if flow1 > 0 or flow2 > 0 or flow3 > 0:
            return "Illumination Alert"
    return ""

def process_excel_file(file_path):
    df = pd.read_excel(file_path)

    df['Message'] = ""

    warning_counts = {'CURRENT WARNING': 0, 'FUSE WARNING': 0, 'VOLTAGE WARNING': 0}

    saved_values = []

    for index, row in df.iterrows():
        saved_values.append(row.tolist())
        print(saved_values)



        if isinstance(saved_values[0][2], float):
            if not math.isnan(saved_values[0][2]):

                Current_Voltage_1 = Error_Detection(float(saved_values[0][2].replace(",", ".")),
                                                    float(saved_values[0][5].replace(",", ".")))
                Current_Voltage_2 = Error_Detection(float(saved_values[0][3].replace(",", ".")),
                                                    float(saved_values[0][6].replace(",", ".")))
                Current_Voltage_3 = Error_Detection(float(saved_values[0][4].replace(",", ".")),
                                                    float(saved_values[0][7].replace(",", ".")))

                if Current_Voltage_1 == "CURRENT WARNING" or Current_Voltage_2 == "CURRENT WARNING" or Current_Voltage_3 == "CURRENT WARNING":
                    message = "CURRENT WARNING"
                    warning_counts["CURRENT WARNING"] += 1

                if Current_Voltage_1 == "VOLTAGE WARNING" or Current_Voltage_2 == "VOLTAGE WARNING" or Current_Voltage_3 == "VOLTAGE WARNING":
                    message = "VOLTAGE WARNING"
                    warning_counts["VOLTAGE WARNING"] += 1

                if Current_Voltage_1 == "FUSE WARNING" or Current_Voltage_2 == "FUSE WARNING" or Current_Voltage_3 == "FUSE WARNING":
                    message = "FUSE WARNING"
                    warning_counts["FUSE WARNING"] += 1

        if isinstance(saved_values[0][2], str):
            if not math.isnan(float(saved_values[0][2].replace(",", "."))):

                Current_Voltage_1 = Error_Detection(float(saved_values[0][2].replace(",", ".")), float(saved_values[0][5].replace(",", ".")))
                Current_Voltage_2 = Error_Detection(float(saved_values[0][3].replace(",", ".")), float(saved_values[0][6].replace(",", ".")))
                Current_Voltage_3 = Error_Detection(float(saved_values[0][4].replace(",", ".")), float(saved_values[0][7].replace(",", ".")))

                if Current_Voltage_1 == "CURRENT WARNING" or Current_Voltage_2 == "CURRENT WARNING" or Current_Voltage_3 == "CURRENT WARNING":
                    message = "CURRENT WARNING"
                    warning_counts["CURRENT WARNING"] += 1

                if Current_Voltage_1 == "VOLTAGE WARNING" or Current_Voltage_2 == "VOLTAGE WARNING" or Current_Voltage_3 == "VOLTAGE WARNING":
                    message = "VOLTAGE WARNING"
                    warning_counts["VOLTAGE WARNING"] += 1

                if Current_Voltage_1 == "FUSE WARNING" or Current_Voltage_2 == "FUSE WARNING" or Current_Voltage_3 == "FUSE WARNING":
                    message = "FUSE WARNING"
                    warning_counts["FUSE WARNING"] += 1

        df.at[index, 'Message'] = message
        message = ""
        Current_Voltage_1 = ""
        Current_Voltage_2 = ""
        Current_Voltage_3 = ""

        saved_values.clear()

    df.to_excel(file_path, index=False)

    show_results(warning_counts)

    show_error_details(df, file_path)

def Illumination_allert_process(file_path, start_hour, end_hour):
    df = pd.read_excel(file_path)
    df['Message'] = ""
    message = ""
    saved_values = []
    warning_count = 0
    for index, row in df.iterrows():
        saved_values.append(row.tolist())
        print(saved_values[0][1][11:16])
        if isinstance(saved_values[0][2], float):
            if not math.isnan(saved_values[0][2]):
                Illumination_Alert = Flow_Detector_With_Time(float(saved_values[0][2].replace(",", ".")),
                                                              float(saved_values[0][3].replace(",", ".")),
                                                              float(saved_values[0][4].replace(",", ".")),
                                                              saved_values[0][1],
                                                              start_hour,
                                                              end_hour)
                if (Illumination_Alert == "Illumination Alert"):
                    message=Illumination_Alert
                    warning_count+=1
                print(warning_count)

        if isinstance(saved_values[0][2], str):
            if not math.isnan(float(saved_values[0][2].replace(",", "."))):
                Illumination_Alert = Flow_Detector_With_Time(float(saved_values[0][2].replace(",", ".")),
                                                              float(saved_values[0][3].replace(",", ".")),
                                                              float(saved_values[0][4].replace(",", ".")),
                                                              saved_values[0][1],
                                                              start_hour,
                                                              end_hour)
                if (Illumination_Alert == "Illumination Alert"):
                    warning_count += 1
                    message = Illumination_Alert
                print(warning_count)

        df.at[index, 'Message'] = message
        message = ""
        Illumination_Alert = ""
        saved_values.clear()
    df.to_excel(file_path, index=False)
    show_results_Illumination(warning_count)
    show_error_details_Illumination(df,file_path)
    warning_count = 0
def Illımunation_alert_process_button():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", ".xlsx;.xls")])
    if file_path:
        # Open a simple dialog to get start and end hours
        start_hour = simpledialog.askstring("Start Hour", "Enter the Start Hour (24-hour format (like 06:30 or 23:30)):")

        end_hour = simpledialog.askstring("End Hour", "Enter the End Hour (24-hour format (like 06:30 or 23:30)):")

        if start_hour is not None and end_hour is not None:
            Illumination_allert_process(file_path, start_hour, end_hour)
def show_results(warning_counts):
    for warning_type, count in warning_counts.items():
        if count > 0:
            messagebox.showwarning(f"{warning_type}", f"{count} instances of {warning_type} detected.")
def show_results_Illumination(warning_count):
        warning_type="Illumination Alert"
        if warning_count > 0:
            messagebox.showwarning( f"{warning_type}",f"{warning_count} illumination warning detected.")
def show_error_details_Illumination(df, file_path):
    error_details = ""

    for index, row in df.iterrows():
        if row['Message']:
            error_details += f"\nRow {index + 1}: {row.iloc[1]} - {row['Message']}"

    if error_details:
        error_details_popup = tk.Tk()
        error_details_popup.title("Error Details")

        text_widget = Text(error_details_popup, wrap=tk.WORD, width=40, height=10)
        text_widget.insert(tk.END, error_details)
        text_widget.config(state=tk.DISABLED)

        scrollbar = Scrollbar(error_details_popup, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Alıcı e-postasını kullanıcıdan al
        alici_email = simpledialog.askstring("Recipient Email", "Enter the Recipient's E-Mail Address:")

        # E-posta gönderme işlemini başlat
        send_email_button = tk.Button(error_details_popup, text="Send Email", command=lambda: send_email(
            "meter.error.detector@gmail.com",           #Gönderen e-posta
            "hgwp zkjv umpp mdgm",          #Gönderen e-posta şifre
            alici_email,            #Alıcı e-posta
            "MED Illumination Data",         #E-posta konusu
            "The Excel spreadsheet has been sent. Good work.",          #E-posta mesajı
            file_path           #E-posta eki (Excel dosyası)
        ))
        send_email_button.pack()

        error_details_popup.mainloop()
def show_error_details(df, file_path):
    error_details = ""

    for index, row in df.iterrows():
        if row['Message']:
            error_details += f"\nRow {index + 1}: {row.iloc[1]} - {row['Message']}"

    if error_details:
        error_details_popup = tk.Tk()
        error_details_popup.title("Error Details")

        text_widget = Text(error_details_popup, wrap=tk.WORD, width=40, height=10)
        text_widget.insert(tk.END, error_details)
        text_widget.config(state=tk.DISABLED)

        scrollbar = Scrollbar(error_details_popup, command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Alıcı e-postasını kullanıcıdan al
        alici_email = simpledialog.askstring("Recipient Email", "Enter the Recipient's E-Mail Address:")

        # E-posta gönderme işlemini başlat
        send_email_button = tk.Button(error_details_popup, text="Send Email", command=lambda: send_email(
            "meter.error.detector@gmail.com",           #Gönderen e-posta
            "hgwp zkjv umpp mdgm",          #Gönderen e-posta şifre
            alici_email,            #Alıcı e-posta
            "MED Current-Voltage Data",         #E-posta konusu
            "The Excel spreadsheet has been sent. Good work.",          #E-posta mesajı
            file_path           #E-posta eki (Excel dosyası)
        ))
        send_email_button.pack()

        error_details_popup.mainloop()

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", ".xlsx;.xls")])
    if file_path:
        process_excel_file(file_path)

app = tk.Tk()
app.iconbitmap(default="bg.ico")
app.title("EnerjiSA - Meter Error Detector")

app.geometry("900x600")
app.resizable(False, False)

bg_image = Image.open("bg.jpg")
bg_image = bg_image.resize((900, 600))
bg_image = ImageTk.PhotoImage(bg_image)

canvas = tk.Canvas(app, width=900, height=600)
canvas.pack()

canvas.create_image(0, 0, anchor=tk.NW, image=bg_image)

button_font = font.Font(size=20, weight="bold")

start_button = tk.Button(app, text="Current-Voltage Detector", command=browse_file, width=20, height=1, bg="navy blue", fg="white", font=button_font)
start_button.pack(pady=50)

start_button.place(relx=0.5, rely=0.8, anchor=tk.CENTER)

akim_hatasi_button = tk.Button(app, text="Illumınatıon Detector", command=Illımunation_alert_process_button, width=20, height=1, bg="navy blue", fg="white", font=button_font)
akim_hatasi_button.pack(pady=50)
akim_hatasi_button.place(relx=0.5, rely=0.9, anchor=tk.CENTER)

app.mainloop()