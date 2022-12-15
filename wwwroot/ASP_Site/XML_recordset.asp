
<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/asp_Site/data/egitim.mdb")
VTNesne.Open ConnectString
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select * from setler;"
KayitSeti.Open SQLKomut, VTNesne , 3 
%>
<?xml version="1.0" encoding="iso-8859-9"?>
<rss version="2.0">
<datadizin>

<%do until kayitseti.EOF %>
  <SET>
    <no><%=KayitSeti.Fields("no")%></no>
    <kategori><%=KayitSeti.Fields("kategori")%></kategori>
    <set_adi><%=KayitSeti.Fields("set_adi")%></set_adi>
    <fiyat><%=KayitSeti.Fields("Fiyat")%></fiyat>
  </SET>
<%
KayitSeti.MoveNext
loop
%>
</datadizin>
</rss>
<%

Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

%>

