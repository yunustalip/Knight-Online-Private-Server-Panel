<%
Set myMail=CreateObject("CDO.Message")
myMail.From="bilgi@yasalegitim.com"
myMail.To="bedriakay@yasalegitim.com"
myMail.CC="bedriakay@yasalegitim.com"
myMail.BCC="bedriakay@yasalegitim.com"
myMail.Subject="CDO ile mesaj göndermek"
myMail.TextBody="Bu CDO ile gönderilmiþtir."
myMail.Send
set myMail=nothing
%>