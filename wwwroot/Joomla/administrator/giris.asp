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
%><head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<meta name="keywords" content="JoomlASP, Joomla, MySQL, ASP, Active Server Page, ASP Portal, JoomlASP temalar�, JoomlASP mod�lleri, JoomlASP bile�enleri, Site i�erik y�netimi, JoomlASP Portal�">
<meta name="description" content="JoomlASP - Geli�ime A��k Site ��erik Y�netimi">
<meta name="author" content="JoomlASP | Hasan Emre Asker">
<title>JoomlASP Site Y�netici Paneli v1.2</title>
<link href="favicon.ico" rel="JoomlASP" />
<link href="admin.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style7 {font-size: 10px}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #990000;
	font-size: 18px;
}
.style9 {color: #990000; font-size: 18px; font-weight: bold;}
.style10 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
-->
</style>
</head>
<!--#include file="../functions/fonksiyonlar.asp"-->
<%
if Session("durum")="giris_yapmis" then
uye_adi = Session("uye_adi")
SQLuye ="SELECT * FROM gop_uyeler where uye_adi='"&uye_adi&"';"
set uye = server.createobject("ADODB.Recordset")
uye.open SQLuye , Baglanti
gid = uye("gid")
if gid = 1 then
Response.Redirect "default.asp"
else
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="../images/admin_top.png"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../images/admin_banner.png" width="307" height="36" /></td>
          <td width="58"><img src="../images/admin_banner_son.png" width="58" height="36" /></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="500"><div align="center">
      <table width="366" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="156"><p><img src="../images/administrator.png" width="128" height="128" /></p>
            <p class="style7">JoomlASP Y�netim Sayfas�na Ho�geldin!</p>
            <p class="style7">&nbsp;</p>
            <p class="style7">L�tfen Administrator kullan�c� ad�n�z� ve �ifrenizi yan alanda girerek giri� yap�n�z.</p>
            <p>&nbsp;</p></td>
          <td width="189"><p class="style9"> Admin Giri�i              </p>
            <p class="style8">&nbsp;</p>
            <form id="form1" name="form1" method="post" action="../default.asp?islem=uye_kontrol">
              <table width="100%" border="0" cellpadding="2" cellspacing="2">
                <tr>
                  <td><span class="style10">�ye Ad�</span></td>
                </tr>
                <tr>
                  <td><input name="uye_adi" type="text" class="inputbox2" id="uye_adi" /></td>
                </tr>
                <tr>
                  <td><span class="style10">�ifre</span></td>
                </tr>
                <tr>
                  <td><input name="uye_sifre" type="password" class="inputbox2" id="uye_sifre" /></td>
                </tr>
                <tr>
                  <td><input name="button" type="submit" class="button" id="button" value="Giri� Yap" /></td>
                </tr>
              </table>
              </form>
            </td>
        </tr>
      </table>
      </div></td>
  </tr>
  <tr>
    <td height="25" background="../images/admin_top2.png"><div align="center" class="style1 style4">JoomlASP Geli�ime A��k Site Y�netimi Sistemi v1.2 </div></td>
  </tr>
</table>
<%
end if
uye.close
set uye=nothing
else
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="../images/admin_top.png"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../images/admin_banner.png" width="307" height="36" /></td>
          <td width="58"><img src="../images/admin_banner_son.png" width="58" height="36" /></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="500"><div align="center">
      <table width="366" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="156"><p><img src="../images/administrator.png" width="128" height="128" /></p>
            <p class="style7">JoomlASP Y�netim Sayfas�na Ho�geldin!</p>
            <p class="style7">&nbsp;</p>
            <p class="style7">L�tfen Administrator kullan�c� ad�n�z� ve �ifrenizi yan alanda girerek giri� yap�n�z.</p>
            <p>&nbsp;</p></td>
          <td width="189"><p class="style9"> Admin Giri�i              </p>
            <p class="style8">&nbsp;</p>
            <form id="form1" name="form1" method="post" action="../default.asp?islem=uye_kontrol">
              <table width="100%" border="0" cellpadding="2" cellspacing="2">
                <tr>
                  <td><span class="style10">�ye Ad�</span></td>
                </tr>
                <tr>
                  <td><input name="uye_adi" type="text" class="inputbox2" id="uye_adi" /></td>
                </tr>
                <tr>
                  <td><span class="style10">�ifre</span></td>
                </tr>
                <tr>
                  <td><input name="uye_sifre" type="password" class="inputbox2" id="uye_sifre" /></td>
                </tr>
                <tr>
                  <td><input name="button" type="submit" class="button" id="button" value="Giri� Yap" /></td>
                </tr>
              </table>
              </form>
            </td>
        </tr>
      </table>
      </div></td>
  </tr>
  <tr>
    <td height="25" background="../images/admin_top2.png"><div align="center" class="style1 style4">JoomlASP Geli�ime A��k Site Y�netimi Sistemi v1.2 </div></td>
  </tr>
</table>
<%
end if
%>