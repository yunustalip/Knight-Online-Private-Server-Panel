<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<%if Session("durum")="esp" Then 
response.expires=0
Session.codepage=1254
response.charset="iso-8859-9"
sir=request.form("siralama")
Response.Write sir
itemler=split(sir,"-")

for i=0 to ubound(itemler)-1
Conne.Execute("update menu set id="&i&" where menuid='"&itemler(i)&"'")
next
End If
%>