<%

Set VTNesne = Server.CreateObject("ADODB.Connection")
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
response.write "Veritaban� kapal�<br>"
end if
Set KayitSeti = Server.CreateObject("ADODB.Recordset")
'KayitSeti.Open "select Set_adi, Fiyat from setler where set_adi LIKE '%HT%';" , VTNesne 
'SQLinsert = "insert into setler (no, kategori, set_adi, fiyat ) values ( 1, 'Web', 'ASP E�i', 50);"
'SQLupdate = "update setler set kategori = 'Sinema' where setler.no = 3;"
KayitSeti.Open "select * from setler where setler.no=3;", VTNesne,,3  
kayitseti("kategori")="Sinema"

On Error Resume Next
KayitSeti.Update
If Err.number = 0 Then
 response.write "Kay�t ba�ar�l� olarak g�ncellendi"
else
 Response.Write "Hata var! Hata Numaras� ve a��klamas� : " 
 Response.write Err.number & " - " & Err.description
End If 

Kayitseti.close
VTNesne.close
Set VTNesne=Nothing

%>
