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
if request.queryString("donus")="sifre" then
Response.Write "<center><br><br><br><br><b>"&notice&":</b><br><br>"&invalid_password&"<br><br><a href='javascript:history.back(1)'><FONT color=#e45f0e><u>"&return&"</u></font></a></b></center>" 
elseif request.QueryString("donus")="uye_adi" then
Response.Write "<center><br><br><br><br><b>"&notice&":</b><br><br>"&invalid_user&"<br><br><a href='javascript:history.back(1)'><FONT color=#e45f0e><u>"&return&"</u></font></a></b></center>" 
elseif request.QueryString("donus")="yok" then
Response.Write "<center><br><br><br><br><b>"&notice&":</b><br><br>"&no_username&"<br><br><a href='javascript:history.back(1)'><FONT color=#e45f0e><u>"&return&"</u></font></a></b></center>" 
end if
%>