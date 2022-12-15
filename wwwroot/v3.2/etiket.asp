<!--#include file="kalip.asp"-->
<% Sub Govde %>
<%
etiket=Trim(Filtre(AramaFiltre(Request.QueryString("etiket"))))
if isnumeric(StrAramaSayi)=false then : StrAramaSayi="5" : end if
%>
<!--#include file="tema/etiket.asp"-->
<% End Sub %>