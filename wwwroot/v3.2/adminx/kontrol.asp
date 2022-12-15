<!--#include file="../ayar.asp"-->
<!--#include file="db.asp"-->
<!--#include file="../inc.asp"-->
<%
response.buffer=true
kullanici=request.form("kullanici")
sifre=request.form("sifre")
sure=request.form("sure")

if kullanici=""&adkull&"" and sifre=""&adsif&"" then
session("admin")="true"
if isnumeric(sure) or sure<>"" then
session.timeout = sure
else
session.timeout = 60
end if
Response.Redirect("default.asp")
Else
response.write("<script>alert('Yanlýþ Kullanýcý Adý veya Parola')</script>")
Response.write("<script>location.href('"& Request.ServerVariables("HTTP_REFERER") &"')</script>")
end if

%>