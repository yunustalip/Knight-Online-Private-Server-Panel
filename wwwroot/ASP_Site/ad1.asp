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
Set DBNesne = Server.CreateObject("ADODB.Connection")
ConnString="dsn=DSN_Egitim;"

'ConnString="Driver={Microsoft Access Driver (*.mdb)};dsn=DSN_Egitim;")
'ConnString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/setler.mdb")
DBNesne.ConnectionString=ConnString
DBNesne.Open 
Set Sonuc = DBNesne.Execute ("create table setler (no int, kategori varchar, set_adi varchar,  fiyat int);")
DBNesne.Close
Set DBNesne=Nothing
%>

