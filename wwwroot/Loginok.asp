<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=8859-9">
<body topmargin="0" marginheight="0">
<%Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Login'")
If MenuAyar("PSt") = 1 Then
Response.Charset = "iso-8859-9"


Dim username,pwd
username = secur(lcase(trim(Request.Form("username"))))
pwd = secur(trim(Request.Form("pwd")))

If Len(username)<=0 Then%>
<table width="200" cellspacing="0" cellpadding="0" border="0" >
<tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">Kullan�c� Giri�i</td></tr>
<tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px"><center>&nbsp;&nbsp;&nbsp;<br>
<b>Kullan�c� Ad�n� Bo� B�rakmay�n�z !</b><br><br>
<a href="javascript:loging()"><b>Geri D�n ve Tekrar Dene</b></a></p>
</font>
</td></tr><tr> <td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>
<%Response.End
ElseIf Len(pwd)<=0 Then%>
<table width="200" cellspacing="0" cellpadding="0" border="0" >
<tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">Kullan�c� Giri�i</td></tr>
<tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px"><center>&nbsp;&nbsp;&nbsp;<br>
<b>�ifre Alan�n� Bo� B�rakmay�n�z !</b><br><br>
<a href="javascript:loging()"><b>Geri D�n ve Tekrar Dene</b></a></p>
</font>
</td></tr><tr> <td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>
<%Response.End
Else
Dim rsUser
Set rsUser = Conne.Execute("Select * From tb_user where strAccountID='"&username&"'")

If not rsUser.eof Then 
Dim rsPwd
Set rsPwd = Conne.Execute("Select strpasswd,strauthority From tb_user where strAccountID='"&username&"' and strPasswd='"&pwd&"'")

If not rsPwd.eof Then
If rspwd("strauthority")="255" Then %>
<table width="200" cellspacing="0" cellpadding="0" border="0" >
<tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">Kullan�c� Giri�i</td></tr>
<tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px">
<center>&nbsp;&nbsp;&nbsp;<br><br><a href="javascript:loging()"><b>Giri�iniz Engellenmi�tir!</b></a><br><br><br>
</td></tr>
<tr><td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>
<%Session("username")=""
Session("login")=""
Session("yetki")=""
Session.abandon
Response.End
End If

Session("login")="ok"
Session("username")=username

Dim ips
ips=Request.ServerVariables("REMOTE_HOST")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&username&" Kullan�c� Giri�i Yapt�.','"&now&"')")
Dim usery
set usery =  Conne.Execute("select * from account_char where straccountid='"&username&"'")
if not usery.eof Then
Dim useryetki
set useryetki = Conne.Execute("select * from USERDATA where struserid='"&usery("strcharid1")&"' or struserid='"&usery("strcharid2")&"' or struserid='"&usery("strcharid3")&"' ")
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
Response.Redirect "login.asp"
Else %>
<table width="200" cellspacing="0" cellpadding="0" border="0" >
<tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">Kullan�c� Giri�i</td></tr>
<tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px"><center>&nbsp;&nbsp;&nbsp;<br>
<b>Kullan�c� Ad� Veya �ifre Hatal� !</b><br><br>
<a href="javascript:loging()"><b>Geri D�n ve Tekrar Dene</b></a></p>
</font>
</td></tr><tr> <td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>
<% End If
else %>
<table width="200" cellspacing="0" cellpadding="0" border="0" >
<tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">Kullan�c� Giri�i</td></tr>
<tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px"><center>&nbsp;&nbsp;&nbsp;<br>
<b>Kullan�c� Ad� Veya �ifre Hatal� !</b><br><br>
<a href="javascript:loging()"><b>Geri D�n ve Tekrar Dene</b></a></p>
</font>
</td></tr><tr> <td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>
<% End If

End If %>
</body>
</html>
<%
MenuAyar.Close
Set MenuAyar=Nothing
else
Response.Write "Login Kapat�lm��t�r."
End If%>