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
'ConnectString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.MapPath("/data/egitim.mdb") & ";Uid=;Pwd=;"
ConnectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/egitim.mdb")
'ConnectString = "Driver={SQL Server};Server=XP1;DataBase=Egitim;Uid=test;Pwd=test;"
'ConnectString = "Driver={MYSQL ODBC 3.51 Driver};Server=localhost;DataBase=Egitim;USER=root;PASSWORD=root;"

'response.write ConnectString & "<br>"
VTNesne.Open ConnectString
if VTNesne.State = 0  then 
'response.write "Veritaban� a��k<br>"           '0.Kapal� 1.A��k 2.A��l�yor
'else
response.write "Veritaban� kapal�<br>"
end if

SQLKomut = "select Set_adi, Fiyat from setler_access where set_adi LIKE '%HT%';"
Set Sonuc=VTNesne.Execute (SQLKomut)
response.write Sonuc("Set_adi")

VTNesne.close
Set VTNesne=Nothing

%>

