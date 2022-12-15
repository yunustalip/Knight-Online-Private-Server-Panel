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
ankid = Request.QueryString ("ankid")
IF Not IsNumeric(Request.QueryString ("ankid")) THEN
response.Redirect "hata.asp"
End if

SQLkat ="SELECT * FROM gop_anakat where ankid=" & guvenlik(request.querystring("ankid")) & ";"
set kat = server.createobject("ADODB.Recordset")
kat.open SQLkat , Baglanti

Response.Write "<div><b>" & kat("ankadi") & "</b><br>" &kat("ankbilgi")&"</div><br><br>"


SQLkat2 ="SELECT * FROM gop_kat where ankid=" & guvenlik(request.querystring("ankid")) & " ;"
set kat2 = server.createobject("ADODB.Recordset")
kat2.open SQLkat2 , Baglanti

listekat2 = 999

for zkat2=1 to listekat2
if kat2.eof then exit for
Response.Write "<b><a href=default.asp?islem=altkategori&katid="&kat2("katid")&">" & kat2("katadi") & "</a></b><br>" &kat2("katbilgi")&"<br><br>"
kat2.MoveNext
next
kat2.close
set kat2=nothing

kat.close
set kat=nothing
%>