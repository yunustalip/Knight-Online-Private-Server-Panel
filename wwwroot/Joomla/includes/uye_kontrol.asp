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
Session.Timeout = 50
If trim(uyeisimkontrol(request.form("uye_adi")))="" then 
Response.Write "<center><br><br><br><br><br><br><br><b>"&notice&":</b><br><br>"&no_username&"<br><br><a href='javascript:history.back(1)'><FONT color=#e45f0e><u>"&return&"</u></a></b></center>" 

else 

If trim(uyeisimkontrol(request.form("uye_sifre")))="" then 
Response.Write "<center><br><br><br><br><br><br><br><b>"&notice&":</b><br><br>"&write_pin&"<br><br><a href='javascript:history.back(1)'><FONT color=#e45f0e><u>"&return&"</u></a></b></center>" 

else

Set rs = Server.CreateObject("Adodb.Recordset") 
Sorgu = "select * from gop_uyeler where uye_adi = '" & uyeisimkontrol(request.form("uye_adi")) & "' and uye_sifre = '" & uyeisimkontrol(md5(Request.form ("uye_sifre"))) & "'" 
rs.Open Sorgu, Baglanti, 1, 3 
If rs.BOF And RS.EOF Then 
Response.Write "<center><br><br><br><br><br><br><br><b>"&notice&":</b><br><br>"&invalid_pass_user&"<br><br><a href='javascript:history.back(1)'><FONT color=#e45f0e><u>"&return&"</u></a></b></center>" 
Else 

Session("durum") = "giris_yapmis" 
Session("uye_id") = rs("uye_id") 
Session("uye_adi") = rs("uye_adi")
Session("gop_language") = rs("uye_dil")
uye_adi = uyeisimkontrol(Request.Form("uye_adi"))
uye_ip = "''" & Request.ServerVariables("REMOTE_ADDR") & "''"
tarih = Year(date)&"-"&Month(date)&"-"&Day(date)&" "&Hour(now)&":"&Minute(now)&":"&second(now)
baglanti.Execute("UPDATE gop_uyeler set uye_son_tarih='"&tarih&"', uye_ip='"&uye_ip&"' where uye_adi='" & uye_adi & "' ;")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If 
end if
end if

%>