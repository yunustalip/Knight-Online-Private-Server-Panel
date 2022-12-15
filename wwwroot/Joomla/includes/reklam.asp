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
set reklam = baglanti.execute("select count(*) from gop_reklam where rgoster = '"& 1 &"'")
k_sayi = reklam(0)
k_sayi = Cint(k_sayi)
set reklam = nothing
Randomize
kac = Int(k_sayi * Rnd)
SQLreklamim = "SELECT * FROM gop_reklam where rgoster = '"& 1 &"';"
SET reklamim = Server.CreateObject("ADODB.Recordset")
reklamim.Open SQLreklamim, Baglanti
if not reklamim.eof then
s = 0
do until reklamim.eof
if s = kac then
hits = reklamim("hit")
baglanti.Execute("UPDATE gop_reklam set hit='"&hits+1&"' where rid=" & reklamim("rid") & " ;")
Response.Write "<center><a href="&reklamim("rlink")&"><img width=""468"" height=""60"" src=" & reklamim("rresim") & " /></a></center>"
end if
s = s + 1
reklamim.movenext
loop
end if
reklamim.close
set reklamim=nothing
%>


