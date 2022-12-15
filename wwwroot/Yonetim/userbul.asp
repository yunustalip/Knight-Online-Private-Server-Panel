<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<%if Session("durum")="esp" Then
if len(request.form("struserid"))>=2 Then
Dim userbul
Set userbul=Conne.Execute("select struserid from userdata where struserid like '%"&request.form("struserid")&"%' order by struserid asc")
If not userbul.eof Then
Do While not userbul.eof
Response.Write "<a href=""javascript:userbul('"&trim(userbul("struserid"))&"');$('#chrs').fadeOut(0);icerikal();"">"&userbul("struserid")&"</a><br>"
Userbul.movenext
Loop
Else
Response.Write "Kullanýcý Bulunamadý"
End If
Else
Response.Write "En Az 2 Karakter Giriniz"
End If
End If%>