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
VTNesne.ConnectionString = "dsn=DSN_Egitim;"
VTNesne.Open '"dsn=DSN_Egitim;"
if VTNesne.State = 0  then 
'response.write "Veritaban� a��k<br>"           '0.Kapal� 1.A��k 2.A��l�yor
'else
response.write "Veritaban� kapal�<br>"
end if
SQLKomut = "select Set_adi, Fiyat from setler_access where set_adi LIKE 'HT%';"
Set Sonuc=VTNesne.Execute (SQLKomut)
response.write Sonuc("Set_adi")
VTNesne.close
Set VTNesne=Nothing

%>

