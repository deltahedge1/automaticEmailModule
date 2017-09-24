class outlook_email(object):
    def __init__(self):
        import win32com.client as win32
        self.outlook = win32.Dispatch("outlook.application")
        pass
        
    def send_mail(self, send_email, subject, body, attachment_path=""):
        try:
            mail = self.outlook.CreateItem(0)
            mail.To=send_email
            mail.Subject = subject
            mail.Body = body
            
            if attachment_path != "":
                mail.Attachments.Add(attachment_path)
       
            mail.send
            
        except:
            pass