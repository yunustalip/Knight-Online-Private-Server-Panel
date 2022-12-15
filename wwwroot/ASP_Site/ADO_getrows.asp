
<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
response.write "Veritabaný kapalý<br>"
end if
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select Set_adi, Fiyat from setler order by set_adi;"
'KayitSeti.Open "select Set_adi, Fiyat from setler where set_adi LIKE '%HT%';" , VTNesne 
KayitSeti.Open SQLKomut, VTNesne , 3 
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
KayitSeti.Movefirst

str = kayitseti.Getstring (,,"$","#")
response.Write(str)

KayitSeti.Movefirst
GetRowsArray = kayitseti.GetRows
for each aa in GetRowsArray
response.Write "-" & aa & "-" &"<br>"
next


Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

%>

