<!--
ASP ve VERITABANI
ADO : ActiveX Data Object
SQL : Structured Query Language

Veritaban� ��lemleri
-----------------------------------
1. Veritaban� nesnesini tan�mlay�n
2. Veritaban� ba�lant�s�n� a��n(DSN veya DSN-less)
3. SQL komutunu icra edin
4. Veritaban� ba�lant�s�n� ve nesneyi kapat�n

-->
<%
Set VTNesne = Server.CreateObject("ADODB.Connection")
'ConnectString = "dsn=DSN_Egitim;" 'DSN Ba�lant�s�
ConnectString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.MapPath("/data/egitim.mdb") & ";Uid=;Pwd=;"
'ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
'ConnectString = "Driver={SQL Server};Server=XP1;DataBase=Egitim;Uid=test;Pwd=test;"
'ConnectString = "Driver={MYSQL ODBC 3.51 Driver};Server=localhost;DataBase=Egitim;USER=root;PASSWORD=root;"

'response.write ConnectString & "<br>"
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
'response.write "Veritaban� a��k<br>"           '0.Kapal� 1.A��k 2.A��l�yor
'else
response.write "Veritaban� kapal�<br>"
end if

SQLcreate  = "create table setler (no int, kategori varchar(20), set_adi varchar(30), fiyat int);"
SQLinsert1 = "insert into setler (no, kategori, set_adi, fiyat ) values ( 1, 'Web', 'ASP E�i', 50);"
SQLinsert2 = "insert into setler (no, kategori, set_adi, fiyat ) values ( 2, 'Web', 'HTML 4.0', 15);"
SQLinsert3 = "insert into setler (no, kategori, set_adi, fiyat ) values ( 3, 'Mul', 'Mult.CD', 20);"
SQLupdate = "update setler set kategori = 'Sinema' where setler.no = 3;"
SQLdelete = "delete from setler where setler.no=3;"
SQLdrop   = "drop table setler;"
SQLselect = "select Set_adi, Fiyat from setler_access where set_adi LIKE '%HT%';"

'VTNesne.Execute (SQLcreate)
'VTNesne.Execute (SQLinsert1)
'VTNesne.Execute (SQLinsert2)
'VTNesne.Execute (SQLinsert3)
'VTNesne.Execute (SQLupdate)
'VTNesne.Execute (SQLdelete)
'VTNesne.Execute (SQLdrop)


VTNesne.close
Set VTNesne=Nothing

%>

