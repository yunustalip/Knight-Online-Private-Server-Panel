<!--#include file="kalip.asp"-->
<% Sub Govde %>
<%
gun=filtre(Request.QueryString("gun"))
ay=filtre(Request.QueryString("ay"))
yil=filtre(Request.QueryString("yil"))
			if isnumeric(gun)=false then
					response.redirect "index.asp"
			elseif isnumeric(ay)=false then
					response.redirect "index.asp"
			elseif isnumeric(yil)=false then
					response.redirect "index.asp"
			end if
Tarih = ay & "/" & gun & "/" & yil
if isnumeric(StrAramaSayi)=false then : StrAramaSayi="5" : end if
%>
<!--#include file="tema/takvim.asp"-->
<% End Sub %>