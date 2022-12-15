<!--#include file="kalip.asp"-->
<% Sub Govde %>
<%
kelime=Request.QueryString("ara")
kelime=Trim(Filtre(AramaFiltre(kelime)))
%>
<!--#include file="tema/ara.asp"-->
<% End Sub %>