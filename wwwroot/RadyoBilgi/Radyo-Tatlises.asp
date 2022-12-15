<%response.charset="utf-8"

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
Dosyaismi = lcase(Request.ServerVariables("Script_Name"))

If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://fmradyodinle.net" or  REFERER_DOMAIN="http://www.fmradyodinle.net" or dosyaismi="/default.asp" or dosyaismi="/404.asp"  Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

If Instr(Request.ServerVariables("ALL_HTTP"),"HTTP_X_REQUESTED_WITH:")>0  or dosyaismi="/default.asp" or dosyaismi="/404.asp" Then
Else
Response.Clear
Response.Write "<a href=""http://www.FmRadyoDinle.net"">www.FmRadyoDinle.net</a>"
Response.End
End If

If IsArray(Application("RadyoBilgi")) Then
RadyoBilgi = Application("RadyoBilgi")
Else
ReDim RadyoBilgi(14,1)
RadyoBilgi(11,0) = Now()-1
RadyoBilgi(11,1) = ""
End If

If DateDiff("s",RadyoBilgi(11,0),Now()) > 45 Then

Const ServerURL = "http://live5.radyotvonline.com:9770/"

Gosterim = True

Private  Function BinaryToString(Binary)
Dim  cl1, cl2, cl3, pl1, pl2, pl3
Dim  L
cl1 = 1
cl2 = 1
cl3 = 1
L = LenB(Binary)
Do  While cl1<=L
pl3 = pl3 &  Chr(AscB(MidB(Binary,cl1,1)))
cl1 = cl1 + 1
cl3 = cl3 + 1
If  cl3>300  Then
pl2 = pl2 & pl3
pl3 = ""
cl3 = 1
cl2 = cl2 + 1
If  cl2>200  Then
pl1 = pl1 & pl2
pl2 = ""
cl2 = 1
End  If
End  If
Loop
BinaryToString = pl1 & pl2 & pl3
End  Function


Err.Clear
On Error Resume Next
	Function GETHTTP(adres) 
	Set objXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	lResolve = 1 * 1000
	lConnect = 1 * 1000
	lSend = 1 * 1000
	lReceive = 1 * 1000
	objXmlHttp.setTimeouts lResolve, lConnect, lSend, lReceive
	objXmlHttp.Open "GET" , adres , false
	objXmlHttp.sEnd 
	GETHTTP = objXmlHttp.ResponseBody
	Set objXmlHttp = Nothing
	End Function

Verimiz = GETHTTP(ServerURL)
Veriler = BinaryToString(Verimiz)

If Err.Number <> 0 Then
	Gosterim = False
End If

If Instr(Veriler,"Server is currently down") > 0 Then
	Gosterim = False
End If

IF Gosterim = True Then

Basla3="Current Song: </font></td><td><font class=default><b>"
Bitir3="</b></td></tr></table><br><table cellpadding=0 cellspacing=0 border=0 width=100"
Itibaren3 = Instr(Veriler,Basla3)

Sarki = Mid(Veriler,Itibaren3+53,len(Veriler))
Ekadar3 = Instr(Sarki,Bitir3)
CalanSarki = Mid(Sarki,1,Ekadar3-1)

Response.Write CalanSarki
RadyoBilgi(11,0) = Now()
RadyoBilgi(11,1) = CalanSarki
Application.Lock
Application("RadyoBilgi") = RadyoBilgi
Application.UnLock
End If

Else
Response.Write RadyoBilgi(11,1)
End If

On Error Resume Next
%>