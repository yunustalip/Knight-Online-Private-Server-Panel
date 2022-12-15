<%
session ("Firmaadi")="Yasal Eðitim"
session ("Firmayil")="01/01/2006"
session ("Firmaadi1")="Yasal Eðitim1"
session ("Firmayil1")="01/01/2005"

response.write session.codepage&"<br>"
response.write session.LCID&"<br>"

session.timeout=1 'dakika
session.codepage=1254
session.LCID=1055
response.write session.codepage&"<br>"
response.write "id:" & session.sessionid&"<br>"
response.write session.LCID&"<br>"
'session.Contents.Remove(4)
'session.Contents.Removeall

for each anahtar in session.Contents
response.write anahtar & ":" & session.Contents(anahtar)& "<br>" 
next
session.Abandon

%>