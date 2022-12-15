<table border="1"><tr>
<% For sayac = 1 To 6 Step 1%>
<% if sayac = 4 then exit for %>
<td><%=sayac%></td>
<%Next%>
</tr></table>

<table border="1"><tr>
<%sayac = 1
while sayac < 5%>
<td><%=sayac%></td>
<%
sayac = sayac + 1
wend%>
</tr></table>
<table border="1"><tr>
<% 
sayac = 10
do while sayac < 5%>
<td><%=sayac%></td>
<%if sayac = 3 then exit do
sayac = sayac + 1
loop%>
</tr></table>
<table border="1"><tr>
<%
sayac = 10
do %>
<td>bir kere <%=sayac%></td>
<%if sayac = 3 then exit do
sayac = sayac + 1
loop while sayac < 5
%></tr></table>
<table border="1"><tr>
<% 
sayac = 1
do until sayac = 5%>
<td><%=sayac%></td>
<%if sayac = 3 then exit do
sayac = sayac + 1
loop%>
</tr></table>
<table border="1"><tr>
<%
sayac = 1
do %>
<td>bir kere <%=sayac%></td>
<%if sayac = 3 then exit do
sayac = sayac + 1
loop until sayac = 5
%></tr></table>


<%
dim mevsim
mevsim=array("ilkbahar","yaz","sonbahar","kış")

For Each elemandegeri in mevsim
'elemandegeri=mevsim(3)
response.Write elemandegeri+"<br>"

next




%>
