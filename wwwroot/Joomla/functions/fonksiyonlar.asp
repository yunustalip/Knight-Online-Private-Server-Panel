<%
'      JoomlASP Site Y�netimi Sistemi (CMS)
'
'      Copyright (C) 2007 Hasan Emre ASKER
'
'      This program is free software; you can redistribute it and/or modify it
'      under the terms of the GNU General Public License as published by the Free
'      Software Foundation; either version 3 of the License, or (at your option)
'      any later version.
'
'      This program is distributed in the hope that it will be useful, but WITHOUT
'      ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
'      FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
'      You should have received a copy of the GNU General Public License along with
'      this library; if not, write to the JoomlASP Asp Yaz�l�m Sistemleri., Kargaz Do�al Gaz Bilgi ��lem M�d�rl���
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anla�mas� Gere�i L�tfen Google Reklam B�l�m�n� Sitenizden kald�rmay�n�z. Bu sizin GOOGLE reklamlar�n� yapman�za
'		kesinlikle bir engel de�ildir. reklam.asp i�eri�inin yada yay�nlad��� verinin de�i�mesi lisans politikas�n�n d���na ��k�lmas�na
'		ve JoomlASP CMS sistemini �cretsiz yay�nlamak yerine �cretlie hale getirmeye bizi te�fik etmektedir. Bu Sistem i�in verilen eme�e
'		sayg� ve bir �e�it �deme se�ene�i olarak GOOGLE reklam�m�z�n de�i�tirmemesi yada silinmemesi gerekmektedir.
%>

<%
sayfayuklenmesi = timer
%>
<!--#include file="db.asp"-->
<%
Set Baglanti= Server.CreateObject("ADODB.Connection")
Baglanti.open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="&mysql_server&"; UID="&mysql_user&"; pwd="&mysql_pass&";db="&mysql_db&"; option = 999999"
	
Ayarlar ="SELECT * FROM gop_ayarlar;"
set ayar = server.createobject("ADODB.Recordset")
ayar.open Ayarlar , Baglanti

'Dil dosyas� ad�n� giriniz uzant�s� olmadan. �rne�in: turkce.asp dili dosyas� i�in dil="turkce" �eklinde yaz�lacakt�r.
'dil = "turkce"
'Dil Ayar� bitmi�tir.

vsayi = ayar("vsayi")
vsutun = ayar("vsutun")
vkarakter = ayar("vkarakter")
surum = "v1.0.3"
vyorum = ayar("vyorum")
siteadi = ayar("siteadi")
session("siteadres") = ayar("siteadres")
site_mail = ayar("admin_mail")
kresimg = ayar("kresim")
google_code = ayar("google_code")
google = ayar("google")
yenileme = ayar("yenileme")
metakey = ayar("meta_key")
metadesc = ayar("meta_desc")
siteadi = ayar("siteadi")
eklenti_k = ayar("eklenti_k")
varsayilandil = ayar("dilayari")

tarih2=Year(date)&"-"&Month(date)&"-"&Day(date)


if Session("durum")="giris_yapmis" then

	Set dil = baglanti.Execute("Select * from gop_uyeler where uye_id=" & session("uye_id") & " ;")
	if dil.eof or dil.bof then 
	Response.write "Hatal� girdi!"
	else

	set language = baglanti.execute("select * from gop_language where lang_id = "& dil("uye_dil") &";")
		if language.eof or language.bof then
		 	set language = baglanti.execute("select * from gop_language where lang_id = "& varsayilandil &";")
			Execute language("language")
		else
		
	Execute language("language")

	dil.close
	set dil=nothing
	end if
	
		end if

else
set language = baglanti.execute("select * from gop_language where lang_id = "& varsayilandil &";")
Execute language("language")
end if

%>
