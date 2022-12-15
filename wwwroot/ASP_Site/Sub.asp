<%
dim par
par=5
Call dongu(par)
%>

<% Sub dongu(parametre)%>
<table border="1"><tr>
<%sayac = 1
while sayac < parametre%>
<td><%=sayac%></td>
<%
sayac = sayac + 1
wend%>
</tr></table>

<%End Sub%>

deneme yapıyoruz : <%call dongu(4)%> yapıldı<br />
deneme yapıyoruz : <%call dongu(8)%> yapıldı<br />