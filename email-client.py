from Tkinter import *
import tkFileDialog
import ttk
import tkMessageBox

import os
import base64
import smtplib
import imaplib
import rfc822

from email import encoders
from email.message import Message
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import email.header
import datetime

root = Tk()
root.title("Email Client Application")

get_path = os.getcwd()

# Global variables, supaya bisa digunakan di banyak fungsi
filename = None
row_browse = None    
namafile = ''
namapath = ''
isSent = 0
isLogin = 0
isAttachment = None
active_burstUpload = 0
namefile_before = ''
attachment = None
mailboxes = ''
num = ''
subject = ''
msg = ''
rv = ''
data = ''
M = None

# Tkinter variables
account = StringVar()
password = StringVar()
receiver = StringVar()
subject = StringVar()
konten = StringVar()
filenamevar = get_path    

def browseFile():
    global isAttachment
    isAttachment = 1
    filepath = tkFileDialog.askopenfilename(title='Choose a file')
    # print "This is the filepath of selected item > " + str(filepath)
    if filepath != None:
        # filename = os.path.split(filepath)[1]
        filename_before = os.path.split(filepath)[0]
        filename = [os.path.join(filename_before, f) for f in os.listdir(filename_before)]
        if filename:
            # print "Filename > " + str(filename)
            print "Folder path > " + str(get_path)
        if filepath:
            print "Filepath > " + str(filepath) + "\n"

        # Define global variables, supaya bisa digunakan di banyak fungsi  
        global namafile 
        global namapath
        global namefile_before
        namafile = filename
        namapath = filepath
        namefile_before = os.path.split(filepath)[1]
        # print namefile_before
        return namafile, namapath, namefile_before
        # encodedContent = base64.b64encode(file)
        # print encodedContent


def clearFile():
    global isAttachment
    isAttachment = 0
    attachment = None
    tkMessageBox.showinfo( "Attachment message", "File Attachment Cleared!\nSudah tidak ada file attachment lagi")
    # print isAttachment
    return isAttachment, attachment


def sendMail():
    global attachment
    global isAttachment

    sender = account.get()
    passwd = password.get()
    recipient = receiver.get()
    subj = subject.get()
    msg = konten.get()
   
    # print "Nama File > " + str(namafile)
    # print "Path File > " + str(namapath)
    # print "Filename before > " + str(namefile_before) 
   
    pesanku = MIMEMultipart()
    pesanku['From'] = sender
    pesanku['To'] = recipient
    pesanku['Subject'] = subj
    pesanku.attach(MIMEText(msg, 'plain'))

    if active_burstUpload == 0:
        if isAttachment == 1:
            attachment = open(namapath, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % namefile_before)
            pesanku.attach(part)
            kontenku = pesanku.as_string()
            print "Message is sent with attachment"
        elif isAttachment == 0:
            kontenku = pesanku.as_string()
            print "Message is sent without attachment"
        else:
            print "Message is sent without attachment"
            kontenku = pesanku.as_string()
            # print "nilai isAttachment di else"
            # print isAttachment

    else:
        print "Message is sent with attachment"
        for file in namafile:
            filename_send = os.path.split(file)[1]
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(open(file, 'rb').read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % filename_send)
            pesanku.attach(part)
            kontenku = pesanku.as_string() 
    
    global isSent
    isSent = 1    

    try:
        mail = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        mail.ehlo()
        mail.login(sender, passwd)
        mail.sendmail(sender, recipient, kontenku)
        mail.close()

        print "Email sent!\n"
        tkMessageBox.showinfo( "Send Email", "Email Sent! " + '\n' + "Sent to: " + str(receiver.get()))
    
    except:
        print "Something went wrong!\n"
        tkMessageBox.showinfo( "Send Email", "Sending Email Failed! " + '\n' + "Sent to : " + str(receiver.get()))


def fetchEmail():
    global mailboxes
    global num
    global subject
    global msg
    global isiMail
    global mainMenuFrame
    global data, rv, result, latest_email_uid
    global emailSubject
    global emailDate
    global fromEmail
    global body

    if rv == 'OK':
            M.list()
            # Out: list of "folders" aka labels in gmail.
            M.select("inbox") # connect to inbox.

            result, data = M.uid('search', None, "ALL") # search and return uids instead
            latest_email_uid = data[0].split()[-1]
            result, data = M.uid('fetch', latest_email_uid, '(RFC822)')
            # result, data = M.uid('fetch', uid, '(X-GM-THRID X-GM-MSGID)')
            raw_email = data[0][1]
            # print raw_email
            isiMail = str(raw_email)
            msg_fetch = email.message_from_string(data[0][1])
            decode_fetch = email.header.decode_header(msg_fetch['Subject'])[0]
            emailSubjectFetch = unicode(decode_fetch[0])
            emailDateFetch = msg_fetch['Date']
            fromEmailFetch = msg_fetch['From']
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True)
                else:
                    continue

    inboxFrame = Toplevel(mainMenuFrame)
    inboxFrame.columnconfigure(0, weight=1)
    inboxFrame.rowconfigure(0, weight=1)

    tkMessageBox.showinfo( "Inbox Message", "Fetching email now...\n" + "email : " + str(sender) + '\n')
    
    ttk.Label(inboxFrame, text="From: ").grid(column=0, row=1, sticky=W)
    ttk.Label(inboxFrame, text="Subject: ").grid(column=0, row=2, sticky=W)
    ttk.Label(inboxFrame, text="Server date: ").grid(column=0, row=3, sticky=W)
    ttk.Label(inboxFrame, text="Message: ").grid(column=0, row=4, sticky=W)

    subjectLabel = ttk.Label(inboxFrame, text=fromEmailFetch)
    subjectLabel.grid(column=1, row=1, padx=2, sticky=W)
    subjectLabel = ttk.Label(inboxFrame, text=emailSubjectFetch)
    subjectLabel.grid(column=1, row=2, padx=2, sticky=W)
    subjectLabel = ttk.Label(inboxFrame, text=emailDateFetch)
    subjectLabel.grid(column=1, row=3, padx=2, sticky=W)

    bodyText = Text(inboxFrame)
    bodyText.grid(column=1, row=4, columnspan=2, rowspan=1, padx=5, sticky=(N, W, E, S))
    bodyText.insert(END, body)

    fullFormat = Toplevel(inboxFrame)
    fullFormat.columnconfigure(0, weight=1)
    fullFormat.rowconfigure(0, weight=1)

    ttk.Label(fullFormat, text="Full Email Format: ").grid(column=0, row=1, sticky=(W, E))

    area = Text(fullFormat)
    area.grid(column=0, row=6, columnspan=2, rowspan=3, padx=5, sticky=(N, W, E, S))
    area.insert(END, isiMail)


def loginMail():
    global account 
    global password 
    global sender
    global receiver 
    global subject 
    global konten 
    global filenamevar
    global rv
    global data
    global M
    global isAttachment
    global isiMail
    global result
    global latest_email_uid

    try :
        global isLogin
        global mailboxes

        sender = account.get()
        passwd = password.get()

        M = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        rv, data = M.login(sender, passwd)
        print rv, data

        if rv == 'OK':
            isLogin = 1
            mainframe.destroy()
            print "Login sukses\n"
            tkMessageBox.showinfo( "Login Message", "Login Sukses!\n" + "Email : " + str(account.get()))

        else:
            tkMessageBox.showinfo( "Login Message", "ERROR: Unable to open mailbox  \n" + rv )
            print "Login gagal\n"

    except:
        isLogin = 0
        tkMessageBox.showinfo( "Login Message", "Login Gagal!\n" + "Email : " + str(account.get()))
        print "Login gagal\n"

    if isLogin == 1:
        global mainMenuFrame
        global filename
        mainMenuFrame = Toplevel(root)
        # mainMenuFrame = ttk.Frame(root, padding="3 3 16 16")
        # mainMenuFrame.grid(column=0, row=0, sticky=(N, W, E, S))
        mainMenuFrame.columnconfigure(0, weight=1)
        mainMenuFrame.rowconfigure(0, weight=1)

        # Sender
        ttk.Label(mainMenuFrame, text="Your Email Account: ").grid(column=0, row=1, sticky=W)
        account.entry = ttk.Entry(mainMenuFrame, width=40, textvariable=account)
        account.entry.config(state='disabled')
        account.entry.grid(column=5, row=1, sticky=(W, E))

        ttk.Label(mainMenuFrame, text="Password: ").grid(column=0, row=2, sticky=W)
        password.entry = ttk.Entry(mainMenuFrame, show="*", width=40, textvariable=password)
        password.entry.config(state='disabled')
        password.entry.grid(column=5, row=2, sticky=(W, E))

        # Receiver
        ttk.Label(mainMenuFrame, text="Recipient's Email: ").grid(column=0, row=3, sticky=W)
        receiver.entry = ttk.Entry(mainMenuFrame, width=40, textvariable=receiver)
        receiver.entry.grid(column=5, row=3, sticky=(W, E))

        # Subject
        ttk.Label(mainMenuFrame, text="Subject: ").grid(column=0, row=4, sticky=W)
        subject.entry = ttk.Entry(mainMenuFrame, width=40, textvariable=subject)
        subject.entry.grid(column=5, row=4, sticky=(W, E))

        # Konten
        ttk.Label(mainMenuFrame, text="Message Body: ").grid(column=0, row=7, ipady=6, sticky=W)
        konten = ttk.Entry(mainMenuFrame, width=40, textvariable=konten)
        konten.grid(column=5, row=7, sticky=(W, E))

        if isAttachment == 1 :
            ttk.Label(mainMenuFrame, text=namefile_before).grid(column=0, row=8, sticky=W)
        
        ttk.Button(mainMenuFrame, text="Browse File", command=lambda:[browseFile()]).grid(column=5, row=8, sticky=W)
        ttk.Button(mainMenuFrame, text="Clear File", command=lambda:[clearFile()]).grid(column=5, row=9, sticky=W)
        ttk.Button(mainMenuFrame, text="Inbox", command=lambda:[fetchEmail()]).grid(column=5, row=10, sticky=W)
        # filenamevar.entry = ttk.Entry(mainframe, width=55, textvariable=filenamevar)
        # filenamevar.entry.grid(column=0, row=8, sticky=(W, E))

        ttk.Button(mainMenuFrame, text="Enable Burst Upload", command=lambda:[enableBurstUpload()]).grid(column=5, row=8, sticky=E)

        if active_burstUpload == 0:
            disable_burstUpload = ttk.Button(mainMenuFrame, text="Disable Burst Upload", command=lambda:[disableBurstUpload()])
            disable_burstUpload.grid(column=5, row=9, sticky=E)
            # disable_burstUpload.pack()

        ttk.Button(mainMenuFrame, text="Send", command=sendMail).grid(column=5, row=10, sticky=E)
        for child in mainMenuFrame.winfo_children(): child.grid_configure(padx=5, pady=5)


def enableBurstUpload():
    global active_burstUpload 
    active_burstUpload = 1
    tkMessageBox.showinfo( "Burst Message", "Burst Upload Activated!\n" + "Upload semua file yang ada dalam folder" )

    return active_burstUpload


def disableBurstUpload():
    global isAttachment
    isAttachment = 0
    attachment = None
    
    global active_burstUpload 
    active_burstUpload = 0
    tkMessageBox.showinfo( "Burst Message", "Burst Upload Deactivated!")

    return active_burstUpload


def aboutUs():
    aboutUsFrame = Toplevel(root)
    aboutUsFrame.geometry("240x200")
    # tkMessageBox.showinfo( "About Message", "Dibuat oleh:\n" + "Rohana Qudus\n" + "Rizaldy Primanta P")
    aboutUsFrame.columnconfigure(0, weight=1)
    aboutUsFrame.rowconfigure(0, weight=1)

    area = Text(aboutUsFrame)
    area.grid(row=1, column=0, columnspan=3, rowspan=4, padx=5, sticky=(N, W, E, S))
    aboutUsKonten = "Dibuat oleh:\n1. Rohana Qudus\n   5115100045\n\n2. Rizaldy Primanta Putra\n   5115100046\n\n3. Hanif Nashrullah\n   5115100140\n\n4. R. Sidqi Tri Priwi\n   5115100153"
    area.insert(END, aboutUsKonten)


def howTo():
    howToFrame = Toplevel(root)
    # tkMessageBox.showinfo( "How to - Instruction", "Cara pakai:\n" + "1. Untuk send hanya dengan 1 attachment file, silahkan isi data pada kolom entry kemudian klik 'Browse File', pilih 1 file dan SEND" + '\n\n' +"2. Untuk send dengan metode Burst Upload, pertama isi data di atas, dan klik tombol Active Burst upload, kemudian pilih direktori / folder anda, kemudian pilih salah satu file, SEMUA file dalam direktori tersebut akan diupload otomatis" + '\n\n' + "Burst upload:\nUpload semua file dalam 1 folder" )
    howToFrame.columnconfigure(0, weight=1)
    howToFrame.rowconfigure(0, weight=1)

    area = Text(howToFrame)
    area.grid(row=1, column=0, columnspan=3, rowspan=4, padx=5, sticky=(N, W, E, S))
    aboutUsKonten = "Cara pakai:\n" + "1. Untuk send hanya dengan 1 attachment file, silahkan isi data pada kolom\n   entry kemudian klik 'Browse File', pilih 1 file dan SEND\n\n2. Untuk send dengan metode Burst Upload, pertama isi data pada kolom entry,\n   klik tombol Enable Burst Upload, lalu pilih direktori/folder anda, kemudian\n   pilih salah satu file, SEMUA file dalam direktori tersebut akan diupload\n   otomatis\n\n3. Tombol clear file berfungsi untuk menghapus file attachment yang dipilih\n   (mengirim tanpa attachment). Tombol clear file hanya berfungsi ketika\n   Burst Upload tidak aktif.\n\nBurst upload:\nUpload semua file dalam 1 folder"
    area.insert(END, aboutUsKonten)


mainframe = ttk.Frame(root, padding="3 3 16 16")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)

aboutmenu = Menu(menubar, tearoff=0)
aboutmenu.add_command(label="How to", command=howTo)
aboutmenu.add_separator()
aboutmenu.add_command(label="About Us", command=aboutUs)
menubar.add_cascade(label="Help", menu=aboutmenu)

# Sender
ttk.Label(mainframe, text="Your Email Account: ").grid(column=0, row=1, sticky=W)
account.entry = ttk.Entry(mainframe, width=40, textvariable=account)
account.entry.grid(column=5, row=1, sticky=(W, E))

ttk.Label(mainframe, text="Password: ").grid(column=0, row=2, sticky=W)
password.entry = ttk.Entry(mainframe, show="*", width=40, textvariable=password)
password.entry.grid(column=5, row=2, sticky=(W, E))

loginbutton = ttk.Button(mainframe, text="Login", command=lambda:[loginMail()]).grid(column=5, row=8, sticky=E)

# Padding for each items
for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)
root.mainloop()