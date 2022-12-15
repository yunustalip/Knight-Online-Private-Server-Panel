<%
dim global
global=6

'donendeger=dongu(global,1)
'response.write(donendeger)
%>

deneme yapıyoruz : <%=dongu(8,7)%> yapıldı<br />

<%=fonkdeger%>


<% Function dongu(carpilacak, carpacak)

dongu=carpilacak * carpacak*global
End Function%>