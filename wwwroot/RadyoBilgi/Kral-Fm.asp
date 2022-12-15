<%Session.codepage=65001
response.charset="utf-8"

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
RadyoBilgi(0,0) = Now()-1
RadyoBilgi(0,1) = ""
End If

If DateDiff("s",RadyoBilgi(0,0),Now()) > 45 Then

Const ServerURL = "http://www.kralfm.com.tr/rds/rds.txt"
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
Function ConvertFromUTF8(sIn)

        Dim oIn: Set oIn = CreateObject("ADODB.Stream")

        oIn.Open
        oIn.CharSet = "iso-8859-9"
        oIn.WriteText sIn
        oIn.Position = 0
        oIn.CharSet = "UTF-8"
        ConvertFromUTF8 = oIn.ReadText
        oIn.Close

End Function

function karakterTR(metin)
metin = replace(metin, "Ã§", "ç")
metin = replace(metin, "Ã‡", "Ç")
metin = replace(metin, "ÄŸ", "ğ")
metin = replace(metin, "Äž", "Ğ")
metin = replace(metin, "Ä±", "ı")
metin = replace(metin, "Ä°", "İ")
metin = replace(metin, "Ã¶", "ö")
metin = replace(metin, "Ã–", "Ö")
metin = replace(metin, "ÅŸ", "ş")
metin = replace(metin, "Åž", "Ş")
metin = replace(metin, "Ã¼", "ü")
metin = replace(metin, "Ãœ", "Ü")
karakterTR = metin
end function



Verimiz = GETHTTP(ServerURL)
Veriler = BinaryToString(Verimiz)

If Err.Number <> 0 Then
	Gosterim = False
End If

IF Gosterim = True Then
Response.Write veriler
RadyoBilgi(0,0) = Now()
RadyoBilgi(0,1) = Veriler
Application.Lock
Application("RadyoBilgi") = RadyoBilgi
Application.UnLock
End If

Else
Response.Write ConvertFromUTF8(RadyoBilgi(0,1))
End If

On Error Resume Next
%>