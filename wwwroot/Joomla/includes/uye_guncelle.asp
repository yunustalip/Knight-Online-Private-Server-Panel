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
if Session("durum")="giris_yapmis" then
uye_id=session("uye_id")

uye_adi=session("uye_adi")
Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr align=""left""><td><div id=""mod_login_username""><b>"&member_info&"</b></div></td></tr></table>"

SQLuye ="SELECT * FROM gop_uyeler where uye_adi='"&uye_adi&"';"
set uye = server.createobject("ADODB.Recordset")
uye.open SQLuye , Baglanti


SQLuye2 ="SELECT * FROM gop_uyeler where uye_id=" & session("uye_id")
set uye2 = server.createobject("ADODB.Recordset")
uye2.open SQLuye2 , Baglanti
uye_id2 = uye2("uye_id")

if uye_id = uye_id2 then
%>

<form name="form1" method="post" action="default.asp?islem=uye_guncelle_islem"><table width=100% border=0 cellpadding=0 cellspacing=0><tr align=left><td width=""20%""><%= username %> </td>  
<td width=""1%"">:</td>  
<td width=""79%""><label>
  <%=uye("uye_adi")%>
</label></td>
</tr>  <tr align=left>    <td><%= name1 %></td>    <td>:</td>    <td><input name="uye_isim" type="text" id="uye_isim" value="<%=uye("uye_isim")%>"></td>  </tr>  
<tr align=left>
  <td><%= name2 %> </td>
  <td>&nbsp;</td>
  <td><input name="uye_soyisim" type="text" id="uye_soyisim" value="<%=uye("uye_soyisim")%>"></td>
</tr>
<tr align=left>
  <td><%= password %></td>
  <td>:</td>
  <td><input name="uye_sifre" type="password" id="uye_sifre" /></td>
</tr>
<tr align=left>    <td><%= email %></td>    <td>:</td>    <td><input name="uye_mail" type="text" id="uye_mail" value="<%=uye("uye_mail")%>"></td>  </tr>  <tr align=left>    <td><%= avatar %></td>    <td>:</td>    <td><input name="uye_avatar" type="text" id="uye_avatar" value="<%=uye("uye_avatar")%>"></td>  </tr>  <tr align=left>    <td><%= web_page %> </td>    <td>:</td>    <td><input name="uye_website" type="text" id="uye_website" value="<%=uye("uye_website")%>"></td>  </tr>  <tr align=left>    <td><%= country %></td>    <td>:</td>    <td><input name="uye_ulke" type="text" id="uye_ulke" value="<%=uye("uye_ulke")%>"></td>  </tr>  <tr align=left>    <td><%= city %></td>    <td>:</td>    <td><input name="uye_sehir" type="text" id="uye_sehir" value="<%=uye("uye_sehir")%>"></td>  </tr>  <tr align=left>    <td><%= msn %></td>    <td>:</td>    <td><input name="uye_msn" type="text" id="uye_msn" value="<%=uye("uye_msn")%>"></td>  </tr>  <tr align=left>    <td><%= icq_number %> </td>    <td>:</td>    <td><input name="uye_icq" type="text" id="uye_icq" value="<%=uye("uye_icq")%>"></td>  </tr>  <tr align=left>    <td><%= aol %></td>    <td>:</td>    <td><input name="uye_aol" type="text" id="uye_aol" value="<%=uye("uye_aol")%>"></td>  </tr>  <tr align=left>    <td><%= yahoo %> </td>    <td>:</td>    <td><input name="uye_yahoo" type="text" id="uye_yahoo" value="<%=uye("uye_yahoo")%>"></td>  </tr>  <tr align=left>    <td><%= skype %></td>    <td>:</td>    <td><input name="uye_skype" type="text" id="uye_skype" value="<%=uye("uye_skype")%>"></td>  </tr>  
<tr align="left">
  <td><%= language %></td>
  <td>:</td>
  <td><select name="uye_dil" id="uye_dil">
  <%
set dil = baglanti.Execute("select * from gop_language order by lang_adi asc")
if dil.eof or dil.bof then
Response.Write "Y�kl� dil yok"
else
do while not dil.eof

if uye("uye_dil") = dil("lang_id") then
Response.Write"<option value="&dil("lang_id")&" selected=""selected"">"&dil("lang_adi")&"</option>"
else
Response.Write "<option value="""&dil("lang_id")&""">"&dil("lang_adi")&"</option>"
end if

dil.movenext
loop
end if
  %>
    </select>
    </td>
</tr>
<tr align=left>    <td>&nbsp;</td>    <td>&nbsp;</td>    <td><label>
  <input type="submit" name="Submit" value="<%= save_change %>">
</label></td>  </tr></table>
</form>
<%
else
Response.Write hello&" "&uye_adi&",<br>"&notice5
end if
uye2.close
set uye2=nothing
uye.close
set uye=nothing
end if
%>
<br>
