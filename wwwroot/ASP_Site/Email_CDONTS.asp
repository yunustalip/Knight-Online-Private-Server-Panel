<%
Set myMail = CreateObject("CDONTS.Newmail")
myMail.From = "bilgi@yasalegitim.com"
myMail.To = "bedriakay@yasalegitim.com"
myMail.CC = ""
myMail.BCC = ""
myMail.Subject = "CDONTS ile mesaj"
myMail.Body = "Bu CDONTS ile gönderilmiþtir."
myMail.Send()
Set myMail = Nothing
%>