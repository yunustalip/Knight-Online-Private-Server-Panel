<!--#include file="../_inc/conn.asp"--><%Response.Charset = "iso-8859-9"
if Session("durum")="esp" Then
itemno=Request.Querystring("itemno")
fmode=Request.Querystring("fmode")
if itemno="" or fmode="" Then
Response.Write "Boþ Býrakmayýnýz."
Response.End
End If
Conne.Execute("exec FindItems "&itemno&","&fmode&"")

set citem=Conne.Execute("select * from ItemFound order by struserid asc,slot asc")
if not citem.eof Then
Response.Write "<table cellspacing=""0"" cellpadding=""3"" width=""400""><tr><td style=""font-weight:bold"">Karakter Adý</td><td  style=""font-weight:bold"">Item No</td><td  style=""font-weight:bold"">Item Serial</td><td style=""font-weight:bold"">Slot</td>"
do while not citem.eof 
Response.Write "<tr onmouseover=""this.style.background='#F5F5F5'"" onmouseout=""this.style.background='#fff'""><td>"&citem(0)&"</td><td>"&citem(1)&"</td><td>"&citem(2)&"</td><td>"&citem(3)&"</td><td><a href=""default.asp?w8=itemfinder&islem=sil&fmode="&fmode&"&struserid="&trim(citem(0))&"&dwid="&citem(1)&"&pos="&citem(3)&""">Sil</a></td></tr>"
citem.movenext
loop

else
Response.Write "Aradýðýnýz Item Kýmsede Bulunamamýþtýr."
End If


else
Response.Redirect("default.asp")
End If
%>