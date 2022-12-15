
<table border="1">
<%
for each var in request.ServerVariables
deger = request.ServerVariables(var) 
%>
<tr><td><%=var%></td><td><%=deger%></td></tr>
<%next%>
</table>

%>
