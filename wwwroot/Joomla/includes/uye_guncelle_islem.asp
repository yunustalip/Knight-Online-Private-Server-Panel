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
if Session("durum")="giris_yapmis" then
uye_adi=session("uye_adi")
uye_id=session("uye_id")

SQLuye ="SELECT * FROM gop_uyeler where uye_adi="& session("uye_id")
set uye = server.createobject("ADODB.Recordset")
uye.open SQLuye , Baglanti

dim oku
Set oku = baglanti.Execute("Select * from gop_uyeler where uye_id='" & session("uye_id") & "' ;")
if oku.eof or oku.bof then
Response.Redirect "404.asp"
else

uye_mail = guvenlik(Request.Form("uye_mail"))
uye_avatar = guvenlik(Request.Form("uye_avatar"))
uye_isim = guvenlik(Request.Form("uye_isim"))
uye_sifre = guvenlik(md5(Request.Form ("uye_sifre")))
uye_soyisim = guvenlik(Request.Form("uye_soyisim"))
uye_website = guvenlik(Request.Form("uye_website"))
uye_ulke = guvenlik(Request.Form("uye_ulke"))
uye_sehir = guvenlik(Request.Form("uye_sehir"))
uye_msn = guvenlik(Request.Form("uye_msn"))
uye_aol = guvenlik(Request.Form("uye_aol"))
uye_icq = guvenlik(Request.Form("uye_icq"))
uye_yahoo = guvenlik(Request.Form("uye_yahoo"))
uye_skype = guvenlik(Request.Form("uye_skype"))
uye_dil = guvenlik(Request.Form("uye_dil"))
gid = guvenlik(oku("gid"))
if guvenlik(Request.Form("uye_sifre")) ="" then
baglanti.Execute("UPDATE gop_uyeler set uye_isim='"&uye_isim&"', uye_soyisim='"&uye_soyisim&"',uye_mail='"&uye_mail&"', uye_avatar='"&uye_avatar&"', uye_website='"&uye_website&"', uye_ulke='"&uye_ulke&"', uye_sehir='"&uye_sehir&"', uye_msn='"&uye_msn&"', uye_icq='"&uye_icq&"', uye_yahoo='"&uye_yahoo&"', uye_aol='"&uye_aol&"', uye_skype='"&uye_skype&"', uye_dil='"&uye_dil&"', gid='"&gid&"' where uye_id='" & uye_id & "' ;")
else
baglanti.Execute("UPDATE gop_uyeler set uye_isim='"&uye_isim&"', uye_sifre='"&uye_sifre&"', uye_soyisim='"&uye_soyisim&"',uye_mail='"&uye_mail&"', uye_avatar='"&uye_avatar&"', uye_website='"&uye_website&"', uye_ulke='"&uye_ulke&"', uye_sehir='"&uye_sehir&"', uye_msn='"&uye_msn&"', uye_icq='"&uye_icq&"', uye_yahoo='"&uye_yahoo&"', uye_aol='"&uye_aol&"', uye_skype='"&uye_skype&"', uye_dil='"&uye_dil&"', gid='"&gid&"' where uye_id='" & uye_id & "' ;")
end if
end if
Response.Write "<center>"&notice6&"</center>"
else
Response.Write hello&" "&uye_adi&",<br>"&notice5

uye.close
set uye=nothing
end if
%>
<br>