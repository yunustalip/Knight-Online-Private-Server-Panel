<!--
Veritabaný Ýþlemleri
 -----------------------------------
1.	Veritabaný nesnesini tanýmlayýn              1. Veritabaný nesnesini tanýmlayýn
2.	Veritabaný baðlantýsýný açýn                 2. Veritabaný baðlantýsýný açýn
3.	Kayýt seti nesnesini oluþturun               3. SQL komutunu icra edin
4.	Kayýt setini açýn                            4. Veritabaný baðlantýsýný ve nesneyi kapatýn
5.	Kayýt setinden istediðiniz kayýtlarý alýn 
6.	Kayýt setini kapatýn 
7.	Baðlantýyý kapatýn ve nesneyi boþaltýn.
-->
<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
response.write "Veritabaný kapalý<br>"
end if
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
SQLKomut = "select * from setler;"
'KayitSeti.Open "select Set_adi, Fiyat from setler where set_adi LIKE '%HT%';" , VTNesne 
KayitSeti.Open SQLKomut, VTNesne , 3 
'KayitSeti.Open server.mappath("Setler.xml"),,,,256    'Kayýtseti.Save ile oluþturulan XML Dosyayý açar
response.write "Okunan Kayýt Sayýsý : " & Kayitseti.RecordCount & "<br>"

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

