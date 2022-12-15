<%
Class VeriOku

Public Aranacak_Kelime
Public Kayit_Sayisi

Sub Kayit_bul

Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString

Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select Set_adi, Fiyat from setler where set_adi LIKE '%" & Aranacak_Kelime& "%';"
KayitSeti.Open SQLKomut, VTNesne , 3 
Kayit_Sayisi = Kayitseti.RecordCount 
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
VTNesne.close
Set VTNesne=Nothing
End Sub

End Class
%>

