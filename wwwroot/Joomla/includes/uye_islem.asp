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
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anla�mas� Gere�i L�tfen Google Reklam B�l�m�n� Sitenizden kald�rmay�n�z. Bu sizin GOOGLE reklamlar�n� yapman�za
'		kesinlikle bir engel de�ildir. reklam.asp i�eri�inin yada yay�nlad��� verinin de�i�mesi lisans politikas�n�n d���na ��k�lmas�na
'		ve JoomlASP CMS sistemini �cretsiz yay�nlamak yerine �cretlie hale getirmeye bizi te�fik etmektedir. Bu Sistem i�in verilen eme�e
'		sayg� ve bir �e�it �deme se�ene�i olarak GOOGLE reklam�m�z�n de�i�tirmemesi yada silinmemesi gerekmektedir.
%>
<%
if session("secure") <> Request.Form("guvenlik_kodu") then
Response.Write "<br><br><center>"&invalid_security_code&"<br><a href=""../default.asp?islem=yeniuye"">"&return&"</a></center>"
else


uye_adi = uyeisimkontrol(Request.Form("uye_adi"))
uye_sifre = md5(Request.Form("uye_sifre"))
uye_sifre2 = md5(Request.Form("uye_sifre2"))
uye_mail = guvenlik(Request.Form("uye_mail"))
uye_isim = guvenlik(Request.Form("uye_isim"))
uye_soyisim = guvenlik(Request.Form("uye_soyisim"))
uye_website = guvenlik(Request.Form("uye_website"))
uye_ulke = guvenlik(Request.Form("uye_ulke"))
uye_sehir = guvenlik(Request.Form("uye_sehir"))
uye_msn = guvenlik(Request.Form("uye_msn"))
uye_icq = guvenlik(Request.Form("uye_icq"))
uye_aol = guvenlik(Request.Form("uye_aol"))
uye_yahoo = guvenlik(Request.Form("uye_yahoo"))
uye_skype = guvenlik(Request.Form("uye_skype"))
gid = "3"
tarih = Year(date)&"-"&Month(date)&"-"&Day(date)&" "&Hour(now)&":"&Minute(now)&":"&second(now)
if uye_adi="" then
Response.Redirect "../default.asp"
else
if uye_sifre=uye_sifre2 then
Baglanti_uye_adi="select * from gop_uyeler where uye_adi='"&uye_adi&"';"
set rs=Baglanti.Execute (Baglanti_uye_adi)
if rs.eof then

SQL2="insert into gop_uyeler (uye_adi,uye_sifre,uye_mail,uye_isim,uye_soyisim,uye_website,uye_ulke,uye_sehir,uye_msn,uye_aol,uye_yahoo,uye_skype,uye_icq,uye_tarih,uye_son_tarih,gid) values ('"&uye_adi&"','"&uye_sifre&"','"&uye_mail&"','"&uye_isim&"','"&uye_soyisim&"','"&uye_website&"','"&uye_ulke&"','"&uye_sehir&"','"&uye_msn&"','"&uye_aol&"','"&uye_yahoo&"','"&uye_skype&"','"&uye_icq&"','"&tarih&"','"&tarih&"','"&gid&"')"
Baglanti.Execute(SQL2)
Response.Write "<center>"&successful_registration&" <br><br><a href=""../default.asp"">"&entry&"</a></center>"

else
Response.Write "<center>"&current_member&"<br><a href=""../default.asp?islem=yeniuye"">"&return&"</a></center>"
end if
else
Response.Write "<center>"&do_not_mach_pass&"<br><a href=""../default.asp?islem=yeniuye"">"&return&"</a></center>"
end if
end if
end if
%>

