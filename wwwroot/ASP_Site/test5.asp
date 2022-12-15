
<%
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
KayitSeti.Open server.mappath("/xml/Setlertxt.xml"),,,,256 

response.write "Okunan Kayýt Sayýsý : " & Kayitseti.RecordCount & "<br>"

%>
<table width="200" border="1">
  <tr>
    <th bgcolor="#FFCCFF" scope="col"><%=KayitSeti.Fields(0).name%></th>
    <th bgcolor="#FFCCFF" scope="col"><%=KayitSeti.Fields(1).name%></th>
  </tr>

<%do until kayitseti.EOF %>
  <tr>
    <td nowrap="nowrap"><%=KayitSeti.Fields("Set_adi")%></td>
    <td nowrap="nowrap"><div align="center"><%=KayitSeti.Fields("Fiyat")%></div></td>
  </tr>
<%
KayitSeti.MoveNext
loop
%>
</table>
<%

Kayitseti.close

%>

