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
SQLinsert = "insert into setler (no, kategori, set_adi, fiyat ) values ( 1, 'Web', 'ASP Eði', 50);"
KayitSeti.Open SQLKomut, VTNesne,,3  

'KayitSeti.AddNew array("no", "kategori", "set_adi", "fiyat" ), array ( 21, "Web", "ASP Eði", 50)
'alanlar  = array("no", "kategori", "set_adi", "fiyat" )
'degerler = array ( 22, "Web", "ASP Eði", 50)

'KayitSeti.AddNew alanlar, degerler

KayitSeti.AddNew
kayitseti("no")=25
kayitseti("kategori")="Deneme"
kayitseti("set_adi")="Deneme Seti"
kayitseti("fiyat")=55
KayitSeti.Update

Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

%>

