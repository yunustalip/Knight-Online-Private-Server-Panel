<script language="javascript" runat="server" src="json2.asp"></script>
<!-- #INCLUDE FILE="fb_app.asp" -->
<style>
.sarkibilgi {

color:#555555;
font: 12px/1.3em Helvetica,Arial,sans-serif;
}
</style><%response.charset="utf-8"
Session.codepage=65001



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
Response.Write "<table class=""sarkibilgi""><tr>"
set liste = xmllist(0).getelementsbytagname("property")

for each i in liste
set a=i.attributes

for each att in a
if att.value = "cue_time_duration" Then
songduration = i.text
End If
if att.value = "track_cover_url" Then
trackcoverurl = i.text
End If
if att.value = "track_artist_name" Then
artist = i.text
End If
if att.value = "track_album_name" Then
album = i.text
End If
if att.value = "cue_title" Then
song = i.text
End If
next

next

sarkim = artist & " - " & song

songmin = songduration/60000
songs= songduration mod 60000

If trackcoverurl<>"" Then
Response.Write "<td rowspan=""7"" class=""text""><img height=""100"" src=""" & replace(trackcoverurl,"albumcover_68x68/","") &"""></td>"& vbcrlf
End If
If artist<>"" Then
Response.Write "<tr><td> Şarkıcı: " & artist &"</td></tr>"& vbcrlf
End If
If album<>"" Then
Response.Write "<tr><td>Albüm: "& Album &"</td></tr>"& vbcrlf
End If
If song<>"" Then
response.Write "<tr><td>Şarkı: "& song &"</td></tr>"& vbcrlf
End If
If songduration<>"" Then
response.Write "<tr><td>Süre: "& int(songmin)&":"&songs/1000 &"</td></tr>"& vbcrlf
End If
Response.Write "<tr><td><a href=""sarki.asp?sarkici="&artist&"&sarki="&song&""" onclick=""SSoz('sarkici="&artist&"&sarki="&song&"');return false"">Şarkı Sözü İçin Tıklayın!</a></td></tr>"
Response.Write "<tr><td>"
%>
<!--#include File="../Mp3.asp"-->
<%
Response.Write "</td></tr>"
Response.Write "</table>"

	%>