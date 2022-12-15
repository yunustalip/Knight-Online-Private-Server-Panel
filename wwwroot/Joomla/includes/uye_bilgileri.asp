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
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<%
if Session("durum")="giris_yapmis" then
uye_id=session("uye_id")
uye_adi=Session("uye_adi")
Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr align=""left""><td><div id=""mod_login_username""><b>"&member_info&"</b></div></td></tr></table>"

SQLuye ="SELECT * FROM gop_uyeler where uye_adi='"&uye_adi&"';"
set uye = server.createobject("ADODB.Recordset")
uye.open SQLuye , Baglanti


Response.Write "<table width=100% border=0 cellpadding=0 cellspacing=0><tr align=left><td width=""20%"">"&username&" </td>  <td width=""1%"">:</td>  <td width=""79%"">"&uye("uye_adi")&"</td></tr>  <tr align=left>    <td>"&name1&" - "&name2&" </td>    <td>:</td>    <td>"&uye("uye_isim") & uye("uye_soyisim")&"</td>  </tr>  <tr align=left>    <td>"&email&"</td>    <td>:</td>    <td>"&uye("uye_mail")&"</td>  </tr>  <tr align=left>    <td>"&avatar&"</td>    <td>:</td>    <td>"&uye("uye_avatar")&"</td>  </tr>  <tr align=left>    <td>"&web_page&" </td>    <td>:</td>    <td>"&uye("uye_website")&"</td>  </tr>  <tr align=left>    <td>"&country&"</td>    <td>:</td>    <td>"&uye("uye_ulke")&"</td>  </tr>  <tr align=left>    <td>"&city&"</td>    <td>:</td>    <td>"&uye("uye_sehir")&"</td>  </tr>  <tr align=left>    <td>"&msn&"</td>    <td>:</td>    <td>"&uye("uye_msn")&"</td>  </tr>  <tr align=left>    <td>"&icq_number&" </td>    <td>:</td>    <td>"&uye("uye_icq")&"</td>  </tr>  <tr align=left>    <td>"&aol&" </td>    <td>:</td>    <td>"&uye("uye_aol")&"</td>  </tr>  <tr align=left>    <td>"&yahoo&" </td>    <td>:</td>    <td>"&uye("uye_yahoo")&"</td>  </tr>  <tr align=left>    <td>"&skype&"</td>    <td>:</td>    <td>"&uye("uye_skype")&"</td>  </tr>  <tr align=left>    <td>&nbsp;</td>    <td>&nbsp;</td>    <td>&nbsp;</td>  </tr></table><br><a href=default.asp?islem=uye_guncelle>"&edit_profil&"</a>"
uye.close
set uye=nothing
else
Response.Write hello&" "&uye_adi&",<br>"& notice5


end if
%>
<br>
