<!--
ASP ve VERITABANI
ADO : ActiveX Data Object
SQL : Structured Query Language

Veritabanı İşlemleri
-----------------------------------
1. Veritabanı nesnesini tanımlayın
2. Veritabanı bağlantısını açın(DSN veya DSN-less)
3. SQL komutunu icra edin
4. Veritabanı bağlantısını ve nesneyi kapatın

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

