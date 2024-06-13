import win32com.client as win32

outlook_app = win32.Dispatch('Outlook.Application')

    # choose sender account
send_account = None
for account in outlook_app.Session.Accounts:
    #print(account.DisplayName)
    #print(account)
    if account.DisplayName == 'email@domail.com':
        send_account = account
        break


adv = '“Subject”'

body = f"""Dear XoXo,

This is the msg {adv}.


Thanks and regards,
Venu.
    """


attachment  = "Path to the file"
    mail = outlook_app.CreateItem(0)

    mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail.To = 'to@mail.com; toto@mail.com'
    mail.CC = 'cc@mail.com; cccc@mail.com'
    
    mail.Subject = f"Something | {adv}"

    mail.Body = body
    mail.Attachments.Add(attachment)

    mail.Send()
