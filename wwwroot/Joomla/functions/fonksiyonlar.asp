<%
'      JoomlASP Site Yönetimi Sistemi (CMS)
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
'      this library; if not, write to the JoomlASP Asp Yazýlým Sistemleri., Kargaz Doðal Gaz Bilgi Ýþlem Müdürlüðü
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
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

'Dil dosyasý adýný giriniz uzantýsý olmadan. Örneðin: turkce.asp dili dosyasý için dil="turkce" þeklinde yazýlacaktýr.
'dil = "turkce"
'Dil Ayarý bitmiþtir.

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
	Response.write "Hatalý girdi!"
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
