import pandas as pd
import datetime
from netmiko import Netmiko
import win32com.client as win32
import os

today = datetime.date.today()


def grab_email():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders['Mal-IPs']
    messages = inbox.items

    for message in messages:
        for attachment in message.attachments:
            yesterday = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d-%m-%y")
            msg_date = message.SentOn.strftime("%d-%m-%y")
            if msg_date == yesterday:
                if str(attachment).startswith("IPs & Domains of Interest"):
                    attachment.SaveAsFile(r'C:\Mal-IPs\mal_upload.xlsx')
                    return message.subject


def upload_to_firewall(file):
    fw_01 = {'host': 'FIREWALL IP',
             'port': PORT,
             'username': 'USER',
             'password': 'PASS',
             'device_type': 'fortinet',
             'fast_cli': True
             }

    print(f"{'#' * 20} Connecting to the Device {'#' * 20} ")
    net_connect = Netmiko(**fw_01)
    count = 0
    count2 = 0
    substring = 'fail'
    failed = 0
    failed2 = 0

    df2 = pd.read_excel(file, 'Malware IP', skiprows=1, skipfooter=1)
    df = pd.read_excel(file, 'Malware Domains', skiprows=1, skipfooter=1)

    for index, row in df.iterrows():
        Domains = row.iloc[0]
        print(Domains)
        config = ['config firewall address',
                  f'edit "pub_mal_{Domains}"',
                  'set type fqdn',
                  f'set fqdn "{Domains}"',
                  'end',
                  'config firewall addrgrp',
                  'edit Bad-IP3',
                  f'append member "pub_mal_{Domains}"',
                  'end'
                  ]
        send_config = net_connect.send_config_set(config)
        print(send_config)
        failed = send_config.count(substring)
        count2 += 1

    for index, row in df2.iterrows():
        IPs = row.iloc[0]
        config = ['config firewall address',
                  f'edit "pub_mal_{IPs}"',
                  'set type ipmask',
                  f'set subnet "{IPs}/32"',
                  'end',
                  'config firewall addrgrp',
                  'edit Bad-IP3',
                  f'append member "pub_mal_{IPs}"',
                  'end'
                  ]
        send_config2 = net_connect.send_config_set(config)
        failed2 = send_config2.count(substring)
        count += 1
    return [count + count2, failed2 + failed]


def send_email(number, failed, subject):
    print("Sending email!")
    email = win32.Dispatch('outlook.application').CreateItem(0)
    emails = 'insert emails here!'

    if number > 0:
        image = r"C:\WalterWhite\danger.png"
        Body = f'<hl>  <font size="+2">The email {subject} sent to *my email* included a total of {number} IPs/Domains, but {failed} failed to upload to the firewall. Group may be full if there are errors, please contact blake for more info </font></hl><br><br><img src="{image}"cid:MyId1"" height=""5"" width=""5""> '

    else:
        image2 = r"C:\WalterWhite\opp1.png"
        Body = f'<hl>  <font size="+2">The email {subject} was sent but included no IPs/Domains to ' \
               f'upload.</font></hl><br><br><img src="{image2}"cid:MyId1"" height=""42"" width=""42""> '

    email.to = emails
    email.Subject = 'Malicious IP/Domain Blocks'
    email.htmlBody = Body
    email.Display()
    email.Send()


def sad_email():
    email = win32.Dispatch('outlook.application').CreateItem(0)
    emails = 'insert emails here!'
    image = r"C:\WalterWhite\wow.png"
    Body = f'<hl>  <font size="+2"> There were no IPs & Domains of Interest Documents forwarded from MS-ISAC Advisory ' \
           f'to *my email* for blocking on the firewall today: {today}  </font></hl><br><br><img src="{image}"cid:MyId1"" height=""5"" width=""5""> '
    email.to = emails
    email.Subject = 'Malicious IP/Domain Blocks'
    email.htmlBody = Body
    email.Display()
    email.Send()


def main():
    subject = grab_email()
    file = 'C:/Mal-IPs/mal_upload.xlsx'
    if os.path.exists(file):
        array = upload_to_firewall(file)
        total = array[0]
        failed = array[1]
        os.remove(file)
        send_email(total, failed, subject)
    else:
        sad_email()


if __name__ == "__main__":
    main()
