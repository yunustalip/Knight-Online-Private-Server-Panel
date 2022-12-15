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
RadyoBilgi(7,0) = Now()-1
RadyoBilgi(7,1) = ""
End If

If DateDiff("s",RadyoBilgi(7,0),Now()) > 45 Then

dim url
url= "http://superfm.karnaval.com/EI/SharedObjects/Handlers/ProxyHandler.ashx?url=http://triton.karnaval.com/mobile_2.xml"
Set xmlObj = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
xmlObj.async = False
xmlObj.setProperty "ServerHTTPRequest", True
xmlObj.Load(url)
If xmlObj.parseError.errorCode <> 0 Then
Response.Write "Bir hata oluştu, RSS kaydı bulunamıyor"
End If
Set xmlList = xmlObj.getelementsbytagname("current_song/song")
Set xmlObj = Nothing
For Each xmlItem In xmlList
Response.Write xmlItem.childNodes(1).text&" - "&xmlItem.childNodes(0).text
RadyoBilgi(7,0) = Now()
RadyoBilgi(7,1) = xmlItem.childNodes(1).text&" - "&xmlItem.childNodes(0).text
Application.Lock
Application("RadyoBilgi") = RadyoBilgi
Application.UnLock
Next

Else
Response.Write RadyoBilgi(7,1)
End If
%>