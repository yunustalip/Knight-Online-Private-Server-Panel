<%
dim global
global=6

'donendeger=dongu(global,1)
'response.write(donendeger)
%>

deneme yap�yoruz : <%=dongu(8,7)%> yap�ld�<br />

<%=fonkdeger%>


<% Function dongu(carpilacak, carpacak)

dongu=carpilacak * carpacak*global
End Function%>