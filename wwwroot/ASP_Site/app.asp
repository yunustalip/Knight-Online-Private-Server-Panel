<%
application ("Firmaadi")="Yasal Eðitim"
application ("Firmayil")="01/01/2006"

for each anahtar in application.Contents
response.write anahtar & ":" & application.Contents(anahtar)& "<br>" 
next

for i=0 to application.Contents.count
response.write application.Contents.key(i) & "<br>"
response.write application.Contents.item(i) & "<br>"
next

response.write application("firmaadi") & "<br>"
response.write application.Contents("firmaadi") & "<br>"


%>