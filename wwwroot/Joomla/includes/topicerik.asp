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
'      this library; if not, write to the JoomlASP Asp Yazılım Sistemleri., Kargaz Doğal Gaz Bilgi İşlem Müdürlüğü
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaşması Gereği Lütfen Google Reklam Bölümünü Sitenizden kaldırmayınız. Bu sizin GOOGLE reklamlarını yapmanıza
'		kesinlikle bir engel değildir. reklam.asp içeriğinin yada yayınladığı verinin değişmesi lisans politikasının dışına çıkılmasına
'		ve JoomlASP CMS sistemini ücretsiz yayınlamak yerine ücretlie hale getirmeye bizi teşfik etmektedir. Bu Sistem için verilen emeğe
'		saygı ve bir çeşit ödeme seçeneği olarak GOOGLE reklamımızın değiştirmemesi yada silinmemesi gerekmektedir.
%>
<h3><%= most_popular %></h3>
  <%
SQLhit ="SELECT * FROM gop_veriler ORDER BY vhit desc;"
set hit = server.createobject("ADODB.Recordset")
hit.open SQLhit , Baglanti

listehit = 10

for zhit=1 to listehit
if hit.eof then exit for
Response.Write "<li class=""mostread""><a href=default.asp?islem=oku&vid="&hit("vid")&" class=""mostread"">" & hit("vbaslik") & "</a></li>"

hit.MoveNext
next
hit.close
set hit=nothing
%>