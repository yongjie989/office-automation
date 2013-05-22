# -*- coding: cp950 -*-

def sendmail_outlook(subject, to, cc="", bcc="", message, attach, ftype, html=False):
    import win32com.client
    import datetime
    f0 = str(datetime.datetime.now())[:10]
    f1 = str(datetime.datetime.now()).split(" ")[-1]
    generated_file_time = f0 + '-'+f1[:2]

    o = win32com.client.Dispatch("Outlook.Application")   
    Msg = o.CreateItem(0)

    # Send to mutiple users can input many email ans separate by ; e.g. user1@example.com;user2@example.com;user3.example.com;
    Msg.To = to

    # CC and BCC user's email
    if cc != "":
        Msg.CC = cc
    if bcc != "":
        Msg.BCC = bcc
    
    Msg.Subject = "%s_%s" % (subject, generated_file_time) 
    if html:
        Msg.HTMLBody = message
    else:
        Msg.Body = message

    # If would like to use html content.
    # Msg.HTMLBody = "<html>..</html>"

    attachment1 = "%s_%s.%s" % (attach, generated_file_time, ftype)
    Msg.Attachments.Add(attachment1)
    # If you have two file need to attach. uncomment lines as below:
    # attachment2 = "Path to attachment no. 2"
    # Msg.Attachments.Add(attachment2)

    Msg.Send()
    return True

sendmail_outlook(
    subject = "",
    to = "Yong Jie Huang <yongjie989@gmail.com>;",
    message = "This is email body...",
    attach = "c:\\automail\\file_you_want_toy_send",
    ftype = "ppt",
    )
    
    