<%
session ("Firmaadi")="Yasal Eðitim"
session ("Firmayil")="01/01/2006"

for each anahtar in session.Contents
response.write anahtar & ":" & session.Contents(anahtar)& "<br>" 
next

for i=0 to session.Contents.count
response.write session.Contents.key(i) & "<br>"
response.write session.Contents.item(i) & "<br>"
next

'response.write session("firmaadi") & "<br>"
'response.write session.Contents("firmaadi") & "<br>"


%>