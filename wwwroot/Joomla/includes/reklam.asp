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


