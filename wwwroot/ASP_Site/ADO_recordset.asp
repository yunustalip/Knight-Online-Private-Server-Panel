<!--
Veritaban� ��lemleri
 -----------------------------------
1.	Veritaban� nesnesini tan�mlay�n              1. Veritaban� nesnesini tan�mlay�n
2.	Veritaban� ba�lant�s�n� a��n                 2. Veritaban� ba�lant�s�n� a��n
3.	Kay�t seti nesnesini olu�turun               3. SQL komutunu icra edin
4.	Kay�t setini a��n                            4. Veritaban� ba�lant�s�n� ve nesneyi kapat�n
5.	Kay�t setinden istedi�iniz kay�tlar� al�n 
6.	Kay�t setini kapat�n 
7.	Ba�lant�y� kapat�n ve nesneyi bo�alt�n.
-->
<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
response.write "Veritaban� kapal�<br>"
end if
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select * from setler;"
'KayitSeti.Open "select Set_adi, Fiyat from setler where set_adi LIKE '%HT%';" , VTNesne 
KayitSeti.Open SQLKomut, VTNesne , 3 
'KayitSeti.Open server.mappath("Setler.xml"),,,,256    'Kay�tseti.Save ile olu�turulan XML Dosyay� a�ar
response.write "Okunan Kay�t Say�s� : " & Kayitseti.RecordCount & "<br>"

'response.write KayitSeti.Fields(0) & "," & KayitSeti.Fields(1) & "<br>"
'response.write KayitSeti.Fields(0).name & "," & KayitSeti.Fields(1).name & "<br>"
'response.write KayitSeti.Fields(0).value & "," & KayitSeti.Fields(1).value & "<br>"
%>
<table width="200" border="1">
  <tr>
    <th bgcolor="#FFCCFF" scope="col"><%=KayitSeti.Fields(0).name%></th>
    <th bgcolor="#FFCCFF" scope="col"><%=KayitSeti.Fields(1).name%></th>
  </tr>

<%do until kayitseti.EOF %>
  <tr>
    <td nowrap="nowrap"><%=KayitSeti.Fields("kategori")%></td>
    <td nowrap="nowrap"><%=KayitSeti.Fields("Set_adi")%></td>
    <td nowrap="nowrap"><div align="center"><%=KayitSeti.Fields("Fiyat")%></div></td>
  </tr>
<%
KayitSeti.MoveNext
loop
%>
</table>
<%

'KayitSeti.save server.mappath("Setler.xml"),1    'XML olarak kaydeder, dosya varsa silmelisiniz
Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

%>

