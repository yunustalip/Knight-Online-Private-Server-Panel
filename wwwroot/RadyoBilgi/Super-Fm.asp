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
RadyoBilgi(1,0) = Now()-1
RadyoBilgi(1,1) = ""
End If

If DateDiff("s",RadyoBilgi(1,0),Now()) > 45 Then



dim url

url= "http://publicapi.streamtheworld.com/public/nowplaying/?mountName=SUPER_FMAAC&numberToFetch=2&eventType=track"

Set xmlObj = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
xmlObj.async = False
xmlObj.setProperty "ServerHTTPRequest", True
xmlObj.Load(url)
If xmlObj.parseError.errorCode <> 0 Then
Response.Write "Bir hata oluştu, RSS kaydı bulunamıyor"
End If
Set xmlList = xmlObj.getelementsbytagname("nowplaying-info")

set liste = xmllist(0).getelementsbytagname("property")

for each i in liste
set a=i.attributes

for each att in a
if att.value = "track_artist_name" Then
artist = i.text
End If
if att.value = "cue_title" Then
song = i.text
End If
next
next

sarkim = artist & " - " & song

Response.Write sarkim
RadyoBilgi(1,0) = Now()
RadyoBilgi(1,1) = sarkim
Application.Lock
Application("RadyoBilgi") = RadyoBilgi
Application.UnLock


Else
Response.Write RadyoBilgi(1,1)
End If
%>