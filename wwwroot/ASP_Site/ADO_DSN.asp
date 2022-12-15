<!--
ASP ve VERITABANI
ADO : ActiveX Data Object
SQL : Structured Query Language

Veritabaný Ýþlemleri
-----------------------------------
1. Veritabaný nesnesini tanýmlayýn
2. Veritabaný baðlantýsýný açýn(DSN veya DSN-less)
3. SQL komutunu icra edin
4. Veritabaný baðlantýsýný ve nesneyi kapatýn

-->
<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
VTNesne.ConnectionString = "dsn=DSN_Egitim;"
VTNesne.Open '"dsn=DSN_Egitim;"
if VTNesne.State = 0  then 
'response.write "Veritabaný açýk<br>"           '0.Kapalý 1.Açýk 2.Açýlýyor
'else
response.write "Veritabaný kapalý<br>"
end if
SQLKomut = "select Set_adi, Fiyat from setler_access where set_adi LIKE 'HT%';"
Set Sonuc=VTNesne.Execute (SQLKomut)
response.write Sonuc("Set_adi")
VTNesne.close
Set VTNesne=Nothing

%>

