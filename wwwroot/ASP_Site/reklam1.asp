<%
Set reklamlar=Server.CreateObject("MSWC.Adrotator")
reklamlar.border=0
'reklamlar.clickable=false
reklamlar.targetframe = "_NEW"
%>
<table border="1"><tr>
<td>bilgi1</td>
<td>www<%=reklamlar.getadvertisement("adrotator.txt")%>zzz</td>
<td>bilgi2</td>
</tr></table>