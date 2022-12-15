<%
application ("Firmaadi")="Yasal Eðitim"
application ("Firmayil")="01/01/2006"
application ("Firmaadi1")="Yasal Eðitim1"
application ("Firmayil1")="01/01/2005"
application.Contents.Removeall(4)
application.Contents.Removeall("firmaadi1")

for each anahtar in application.Contents
response.write anahtar & ":" & application.Contents(anahtar)& "<br>" 
next


%>