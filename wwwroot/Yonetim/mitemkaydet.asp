<% if Session("durum")="esp" Then %>
<!--#include file="../_inc/conn.asp"-->
<%
ssid=Request.Querystring("ssid")
slot=Request.Querystring("slot")
num=Request.Querystring("num")
dropyuzde=Request.Querystring("dropyuzde")
dropyuzde=dropyuzde*100
Conne.Execute("update k_monster_item set iItem0"&slot&"="&num&", sPersent0"&slot&"="&dropyuzde&" where sIndex="&ssid&" ")
End If

%>