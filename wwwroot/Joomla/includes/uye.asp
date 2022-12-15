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
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaşması Gereği Lütfen Google Reklam Bölümünü Sitenizden kaldırmayınız. Bu sizin GOOGLE reklamlarını yapmanıza
'		kesinlikle bir engel değildir. reklam.asp içeriğinin yada yayınladığı verinin değişmesi lisans politikasının dışına çıkılmasına
'		ve JoomlASP CMS sistemini ücretsiz yayınlamak yerine ücretlie hale getirmeye bizi teşfik etmektedir. Bu Sistem için verilen emeğe
'		saygı ve bir çeşit ödeme seçeneği olarak GOOGLE reklamımızın değiştirmemesi yada silinmemesi gerekmektedir.
%>
<%
if Session("durum")="giris_yapmis" then
uye_id=request.querystring("uye_id")

SQLuye ="SELECT * FROM gop_uyeler where uye_id='"&uye_id&"';"
set uye = server.createobject("ADODB.Recordset")
uye.open SQLuye , Baglanti


Response.Write "<table width=100% border=0 cellpadding=0 cellspacing=0><tr align=left><td width=""20%"">"&username&" </td>  <td width=""1%"">:</td>  <td width=""79%"">"&uye("uye_adi")&"</td></tr>  <tr align=left>    <td>"&name1&" - "&name2&" </td>    <td>:</td>    <td>"&uye("uye_isim") & uye("uye_soyisim")&"</td>  </tr>  <tr align=left>    <td>"&web_page&" </td>    <td>:</td>    <td>"&uye("uye_website")&"</td>  </tr>  <tr align=left>    <td>"&country&"</td>    <td>:</td>    <td>"&uye("uye_ulke")&"</td>  </tr>  <tr align=left>    <td>"&city&"</td>    <td>:</td>    <td>"&uye("uye_sehir")&"</td>  </tr>  <tr align=left>    <td>"&msn&"</td>    <td>:</td>    <td>"&uye("uye_msn")&"</td>  </tr>  <tr align=left>    <td>"&icq_number&" </td>    <td>:</td>    <td>"&uye("uye_icq")&"</td>  </tr>  <tr align=left>    <td>"&aol&"</td>    <td>:</td>    <td>"&uye("uye_aol")&"</td>  </tr>  <tr align=left>    <td>"&yahoo&" </td>    <td>:</td>    <td>"&uye("uye_yahoo")&"</td>  </tr>  <tr align=left>    <td>"&skype&"</td>    <td>:</td>    <td>"&uye("uye_skype")&"</td>  </tr>  <tr align=left>    <td>&nbsp;</td>    <td>&nbsp;</td>    <td>&nbsp;</td>  </tr></table>"
%>

<form id="form1" name="form1" method="post" action="?islem=mesaj_gonder&auid=<%= uye("uye_id")%>">
  <table width="100%" border="0" cellpadding="2" cellspacing="2">
    <tr>
      <td height="25" colspan="3"><div align="center"><%= send_message %></div></td>
    </tr>
    <tr>
      <td width="73"><%= heading %></td>
      <td width="5">:</td>
      <td width="921"><input type="text" name="mesaj_baslik" id="mesaj_baslik" class="inputbox" /></td>
    </tr>
    <tr>
      <td><%= message %></td>
      <td>:</td>
      <td><textarea name="mesaj_icerik" id="mesaj_icerik" cols="45" rows="5" class="inputbox"></textarea></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><input type="submit" name="button" id="button" value="<%= send_message %>" class="button" /></td>
    </tr>
  </table>
</form>
<%
uye.close
set uye=nothing

else
Response.Write "<center>" & notice4 & "</center>"
end if
%>
<br>
