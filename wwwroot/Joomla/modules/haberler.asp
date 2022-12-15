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
<table width="100%" border="0" cellpadding="2" cellspacing="2"><tr align="left">
<%
SQLveri ="SELECT * FROM gop_veriler where katid ='"& 1 &"' order by vid desc;"
set veri = server.createobject("ADODB.Recordset")
veri.open SQLveri , Baglanti

deste = 999

for z=1 to deste
if veri.eof then exit for
SQLyazar ="SELECT * FROM gop_uyeler where uye_id ='"& veri("uye_id") &"';"
set yazar = server.createobject("ADODB.Recordset")
yazar.open SQLyazar , Baglanti
Response.Write "<td width=""50%"" valign=""top""><p><b>Yazan:</b>"&yazar("uye_adi")&"<br>"&veri("vtarih")&"</p><br>"
yazar.close
set yazar=nothing
Response.Write "<b><img src="& veri("vresim") &" align=""left"" onerror=""this.src='images/joomlasp.jpg';"" align=left>" & veri("vbaslik") & "...</b><br>"&left(veri("vicerik"),200)&"... <a href=default.asp?islem=oku&vid="&veri("vid")&">Devamý...</a><br>Görüntülenme sayýsý: "&veri("vhit")&"</td>"
if z mod 2 = 0 then response.write "</tr><tr>"
veri.MoveNext
next
veri.close
set veri=nothing
%>
</tr></table>