<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
response.write "Veritabaný kapalý<br>"
end if
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
'KayitSeti.Open "select Set_adi, Fiyat from setler where set_adi LIKE '%HT%';" , VTNesne 
'SQLinsert = "insert into setler (no, kategori, set_adi, fiyat ) values ( 1, 'Web', 'ASP Eði', 50);"
'SQLupdate = "update setler set kategori = 'Sinema' where setler.no = 3;"
KayitSeti.Open "select * from setler where setler.no>=20;", VTNesne,3,3  
'response.Write "sayý:" & kayitseti.recordcount
'if kayitseti.recordcount = 1 then KayitSeti.Delete
do until kayitseti.eof
KayitSeti.Delete
kayitseti.movenext
loop


Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

%>

