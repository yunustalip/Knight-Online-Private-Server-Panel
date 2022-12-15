<%

Set objNewMail = CreateObject("CDONTS.Newmail")
objNewMail.From = "bilgi@yasalegitim.com"
objNewMail.To = "bedriakay@yasalegitim.com"
objNewMail.CC = ""
objNewMail.Subject = "Sending email with CDO"
objNewMail.Body = "This is a message."

objNewMail.Send()

Set objNewMail = Nothing
%>