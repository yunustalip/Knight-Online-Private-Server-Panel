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

Const ServerURL = "http://www.radyo5.com.tr/calansarki.asp"
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

Basla3="<marquee class=""baslik"" scrolldelay=""150"" OnMouseOver=""this.stop()"" OnMouseOut=""this.start()"">"
Bitir3="</marquee>"
Itibaren3 = Instr(Veriler,Basla3)

Sarki = Mid(Veriler,Itibaren3+len(Basla3),len(Veriler))
Ekadar3 = Instr(Sarki,Bitir3)
CalanSarki = Mid(Sarki,1,Ekadar3-1)
End If

IF Gosterim = True Then
Response.Write CalanSarki
End If

On Error Resume Next
%>