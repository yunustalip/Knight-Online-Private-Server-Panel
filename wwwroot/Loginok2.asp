<!--#include file="_inc/conn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=8859-9">
<body topmargin="0" marginheight="0">
<%Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Login'")
If MenuAyar("PSt")=1 Then
Response.Charset = "iso-8859-9"
function guvenlik(data) 
Data = Replace( data , "'" , "", 1, -1,1)
data = Replace (data ,"`","",1,-1,1) 
data = Replace (data ,"=","",1,-1,1) 
data = Replace (data ,"&","",1,-1,1) 
data = Replace (data ,"%","",1,-1,1) 
data = Replace (data ,"!","",1,-1,1) 
data = Replace (data ,"#","",1,-1,1) 
data = Replace (data ,"<","",1,-1,1) 
data = Replace (data ,">","",1,-1,1) 
data = Replace (data ,"*","",1,-1,1) 
data = Replace (data ,",","",1,-1,1) 
data = Replace (data ,"'","",1,-1,1) 
data = Replace (data ,"Chr(34)","",1,-1,1) 
data = Replace (data ,"Chr(39)","",1,-1,1) 
guvenlik=data 
end function 

dim username,pwd
username = guvenlik(lcase(trim(Request.Form("username"))))
pwd = guvenlik(trim(Request.Form("pwd")))

if username="" Then
with response
.write "<font face=""arial,helvetica"" size=""2"">"
.write "<p align=""center""><b>Kullanýcý adýný boþ býrakmayýnýz.</b><br><br>"
.write "<a href=""hata.asp""><b>Önceki sayfaya dön</b></a></p>"
.write "</font>"
end with
Response.End
elseif pwd="" Then
with response
.write "<font face=""arial,helvetica"" size=""2"">"
.write "<p align=""center""><b>Þifre alanýný boþ býrakmayýnýz.</b><br><br>"
.write "<a href=""hata.asp""><b>Önceki sayfaya dön</b></a></p>"
.write "</font>"
end with
Response.End
else
dim rsUser,sql
Set rsUser = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&username&"'"
rsUser.open sql,conne,1,3

if not rsUser.eof Then 
Dim rsPwd
Set rsPwd = Server.CreateObject("ADODB.Recordset")
sql = "Select * From tb_user where strAccountID='"&username&"' and strPasswd='"&pwd&"'"
rsPwd.open sql,conne,1,3

if not rsPwd.eof Then
if rspwd("strauthority")="255" Then
response.write "<br><br><br><b>Giriþiniz Yasaklanmýþtýr.</b>"
Session("username")=""
Session("login")=""
Session("yetki")=""
Session.abandon
Response.write "<script>alert('Hesabýnýz Bloke Olmuþtur. Admin Ýle Ýletiþime Geçiniz');self.close()</script>"
Response.End
End If
Session("login")="ok"
Session("username")=username
Dim usery
Set Usery =  Conne.Execute("select * from account_char where straccountid='"&username&"'")
If not usery.eof Then
Dim useryetki
Set useryetki = Conne.Execute("select * from USERDATA where struserid='"&usery("strcharid1")&"' or struserid='"&usery("strcharid2")&"' or struserid='"&usery("strcharid3")&"' ")
If Not Useryetki.Eof Then
Do While Not UserYetki.Eof
If Useryetki("Authority")="0" Then
Session("yetki")="1"
Exit Do
Else
Session("yetki")=""
End If
UserYetki.MoveNext
Loop
End If

End If
Response.Redirect("powerupstore.asp")
Else %>

<font face="arial,helvetica" size="2">
<p align="center"><b>Þifre Yanlýþ !</b><br><br>
<a href="hata.asp"><b>Geri Dön ve Tekrar Dene</b></a></p>
</font>

<% End If
else %>
<font face="arial,helvetica" size="2">
<p align="center"><strong>Kullanýcý Adý Hatalý</strong><br><br>
<a href="hata.asp"><b>Geri Dön ve Tekrar Dene</b></a></p>
</font>

<% End If

End If %>
</body>
</html>
<%
MenuAyar.Close
Set MenuAyar=Nothing
else
Response.Write "Login Kapatýlmýþtýr."
End If%>