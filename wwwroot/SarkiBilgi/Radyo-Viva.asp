﻿<%
response.expires=-1
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

url= "http://www.showradyo.com.tr/rds/vivasongdata.xml"
Set createxml = Server.CreateObject("msxml2.DOMDocument")
createxml.async = false
createxml.SetProperty "ServerHTTPRequest", True
createxml.Load(url)
Set xmllist = createxml.getElementsByTagName("rds")

for each xmlveri in xmllist
start=xmlveri.childNodes(0).attributes.getNamedItem("start").nodeValue
sure=xmlveri.childNodes(0).attributes.getNamedItem("duration").nodeValue
artist=xmlveri.childNodes(0).attributes.getNamedItem("artist").nodeValue
sarki=xmlveri.childNodes(0).attributes.getNamedItem("title").nodeValue
album=xmlveri.childNodes(0).attributes.getNamedItem("album").nodeValue
cover=xmlveri.childNodes(0).attributes.getNamedItem("cover").nodeValue
next

if start<>"" and artist<>"" and sarki<>"" and sure<>"" and album<>"" and cover<>"" Then
%>
<style>
.sarkibilgi {

color:#555555;
font: 12px/1.3em Helvetica,Arial,sans-serif;
}
.sarkibilgi a{
color:#FFFFFF;
text-decoration:none;
}
.sarkibilgi a:hover{
	color: #ffffff;
	text-decoration:underline;
}

</style>
<table border="0" cellspacing="0" class="sarkibilgi">
<tr>
<td rowspan="4"><img src="SarkiBilgi/RadyoVivaCover.Asp" height="85"  width="85" onError=src="cover.png"></td>
<td>Sanatçı</td>
<td>:</td>
<td><%=artist%></td>
</tr>

<tr>
<td>Albüm</td>
<td>:</td>
<td><%=album%></td>
</tr>
<tr>

<td>Şarkı</td>
<td>:</td>
<td><%=sarki%></td>
</tr>

<tr>
<td>Şarkı Süresi</td>
<td>:</td>
<td><%dk=int(sure/60)
sn=sure mod 60
if len(dk)=1 Then
dk="0"&dk
End If
if len(sn)=1 Then
sn="0"&sn
End If
Response.Write dk&":"&sn

sarkim = artist & " - " & "sarki"
%></td>
</tr>
<tr>
<td>
<!--#include File="../Mp3.asp"-->
</td>
</tr>
</table><%
Function Base64Decode(ByVal base64String)
Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim dataLength, sOut, groupBegin
 
base64String = Replace(base64String, vbCrLf, "")
base64String = Replace(base64String, vbTab, "")
base64String = Replace(base64String, " ", "")

dataLength = Len(base64String)
If dataLength Mod 4 <> 0 Then
Err.Raise 1, "Base64Decode", "Bad Base64 string."
Exit Function
End If

For groupBegin = 1 To dataLength Step 4
Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
numDataBytes = 3
nGroup = 0

For CharCounter = 0 To 3

thisChar = Mid(base64String, groupBegin + CharCounter, 1)

If thisChar = "=" Then
numDataBytes = numDataBytes - 1
thisData = 0
Else
thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
End If
If thisData = -1 Then
Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
Exit Function
End If

nGroup = 64 * nGroup + thisData
Next

nGroup = Hex(nGroup)

nGroup = String(6 - Len(nGroup), "0") & nGroup

pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
Chr(CByte("&H" & Mid(nGroup, 5, 2)))

sOut = sOut & Left(pOut, numDataBytes)
Next

Base64Decode = sOut
End Function

Function StringToMultiByte(S)
Dim i, MultiByte
For i=1 To Len(S)
MultiByte = MultiByte & ChrB(Asc(Mid(S,i,1)))
Next
StringToMultiByte = MultiByte
End Function

if len(cover)>0 Then
kod=StringToMultiByte(Base64Decode(cover))
else
kod=""
End If

if artist="RADYO VİVA" or artist="" Then
else
artist=replace(artist,"'","&#39;")
artist=replace(artist,"&","&#38;")
sarki=replace(sarki,"'","&#39;")
sarki=replace(sarki,"&","&#38;")
album=replace(album,"'","&#39;")
album=replace(album,"&","&#38;")



End If

End If
%>