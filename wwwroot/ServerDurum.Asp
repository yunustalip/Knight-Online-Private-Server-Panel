<!--#include file="_inc/conn.asp"-->
<%Session.codepage=1254
Response.Expires=0
Response.Charset="iso-8859-9"
Dim sitesettings,IP
Set sitesettings=Conne.Execute("select IP from siteayar")
IP = sitesettings("IP")

Dim LoginServer,GameServer,FtpServer,LoginServerDurum,GameServerDurum,FtpServerDurum
Dim OnlineUserCount
Set OnlineUserCount = Conne.Execute("Select Count(strcharid) as ToplamOnline From CurrentUser")

Set LoginServer = Server.CreateObject("Port.CSocket")
LoginServer.connect IP, 15100
Set GameServer = Server.CreateObject("Port.CSocket")
GameServer.connect IP, 15001
Set FtpServer = Server.CreateObject("Port.CSocket")
FtpServer.connect IP, 21
If LoginServer.State=7 Then
LoginServerDurum="<img src=""imgs/acc.gif"" align=""absmiddle"">"
Else
LoginServerDurum="<img src=""imgs/rej.gif"" align=""absmiddle"">"
End If
If GameServer.State=7 Then
GameServerDurum="<img src=""imgs/acc.gif"" align=""absmiddle"">"
Else
GameServerDurum="<img src=""imgs/rej.gif"" align=""absmiddle"">"
End If
If FtpServer.State=7 Then
FtpServerDurum="<img src=""imgs/acc.gif"" align=""absmiddle"">"
Else
FtpServerDurum="<img src=""imgs/rej.gif"" align=""absmiddle"">"
End If

If Request.Querystring("islem")=2 Then
With Response
	.Write "<strong>Login Server : </strong>"&LoginServerDurum
	.Write "<br>"
	.Write "<strong>Game Server : </strong>"&GameServerDurum
	.Write "<br>"
	.Write "<strong>FTP Server : </strong>"&FtpServerDurum
	.Write "<br>"
	End With
Response.End
End If
Response.Write "Login Server: "&LoginServerDurum&"&nbsp;&nbsp;Game Server: "&GameServerDurum&"&nbsp;&nbsp;FTP Server: "&FtpServerDurum&"&nbsp;&nbsp;Online Oyuncu: " & OnlineUserCount("ToplamOnline")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<code style="color:#00ff00;font-weight:bold;font-size:11px">Güvenlik Sistemi Devrede! Flood sayýnýz (<%
if Session("FloodSayi")="" Then
Response.Write "0"
Else
Response.Write Session("FloodSayi")
End If%>/6)</code> <span style="text-decoration:blink;color:#00ff00;">_</span>