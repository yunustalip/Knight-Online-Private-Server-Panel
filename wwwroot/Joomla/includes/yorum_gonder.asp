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
<% if Session("durum")="giris_yapmis" then %>
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>

<%
uye_id=session("uye_id")

yorum = guvenlikyorum(Request.Form("yorum"))
uye_id = guvenlik(uye("uye_id"))
vid = guvenlik(request.querystring("vid"))
yorum_tarih = tarih2
yorum_onay = "0"

SQL="insert into gop_yorumlar (yorum, uye_id, vid, yorum_tarih, yorum_onay) values ('"&yorum&"','"&uye_id&"','"&vid&"','"&yorum_tarih&"','"&yorum_onay&"')"
Baglanti.Execute(SQL)
Response.Write "<center>"&thanks&"<br>"&thanks2&"</center>"

else
Response.write notice4
end if 
%>
