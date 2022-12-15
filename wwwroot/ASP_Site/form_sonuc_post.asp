TotalBytes : <%=request.TotalBytes%><br>

Adý Soyadý : <%=request.form("ADSOYAD")%><br>
E-Posta : <%=request.form("email")%><br>
Buton : <%=request.form("buton")%><br>
gizli : <%=request.form("gizlibilgi")%><br>
<%= "---------------<br>"%>
radyo : <%=request.form("Radyo")%><br>
SecimListesi : <%=request.form("SecimListesi")%><br>

<%
for each degisken in request.Form("isaretkutusu")
response.Write degisken & "<br>"
next
%>
<%
for each degisken in request.Form("CokluSecim")
response.Write degisken & "<br>"
next
%>

<%= "---------------<br>"%>


<%
for each degisken in request.Form
response.Write degisken & ":" & request.form(degisken) & "<br>"
next
%>