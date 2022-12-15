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
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
.style2 {color: #FF0000}
-->
</style>

<script language="javascript" type="text/javascript" src="form.js"></script>
<form name="form1" method="post" action="default.asp?islem=uye_islem">
  <table width="90%" border="0" align="center" cellpadding="2" cellspacing="2" height="685">
    <tr>
      <td colspan="3" height="21"></td>
    </tr>
    <tr>
      <td colspan="3" height="21" background="ust_bg.jpg"><span class="style1"><%= member_info %></span></td>
    </tr>
    <tr>
      <td colspan="3" height="21"><span class="style1"></span></td>
    </tr>
    <tr>
      <td height="25"><span class="style2">* </span><%= username %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_adi" type="text" id="uye_adi"
      maxlength="25" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><span class="style2">* </span><%= password %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_sifre" type="password" id="uye_sifre"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><span class="style2">* </span><%= confirm_password %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_sifre2" type="password" id="uye_sifre2"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><span class="style2">* </span><%= email %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_mail" type="text" id="uye_mail"
      maxlength="35" / size="20"></td>
    </tr>
    <tr>
      <td height="21"></td>
      <td height="21"></td>
      <td height="21"></td>
    </tr>
    <tr>
      <td colspan="3" height="21" background="ust_bg.jpg"><span class="style1"><%= personel_data %></span></td>
    </tr>
    <tr>
      <td colspan="3" height="21"><span class="style1"></span></td>
    </tr>
    <tr>
      <td height="25"><%= name1 %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_isim" type="text" id="uye_isim"
      maxlength="25" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= name2 %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_soyisim" type="text" id="uye_soyisim"
      maxlength="25" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= web_page %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_website" type="text" id="uye_website"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= country %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_ulke" type="text" id="uye_ulke"
      maxlength="25" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= city %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_sehir" type="text" id="uye_sehir"
      maxlength="25" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= msn %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_msn" type="text" id="uye_msn"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= icq_number %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_icq" type="text" id="uye_icq"
      maxlength="15" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= aol %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_aol" type="text" id="uye_aol"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= yahoo %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_yahoo" type="text" id="uye_yahoo"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="25"><%= skype %></td>
      <td height="25">:</td>
      <td height="25"><input name="uye_skype" type="text" id="uye_skype"
      maxlength="50" / size="20"></td>
    </tr>
    <tr>
      <td height="21"><%= security_code %></td>
      <td height="21">:</td>
      <td height="21"><%
session("secure") = secure_code
Response.Write session("secure")

 %></td>
    </tr>
    <tr>
      <td height="21"></td>
      <td height="21"></td>
      <td height="21"><input type="text" name="guvenlik_kodu" id="guvenlik_kodu" /></td>
    </tr>
    <tr>
      <td height="21"></td>
      <td height="21"></td>
      <td height="21"></td>
    </tr>
    <tr>
      <td colspan="3" height="27"><dl>
        <dd align="center"><input name="Submit2" type="submit" class="buton" value="<%= complete_registration %>"> <input
          type="reset" name="button2" id="button2" value="<%= reset %>" /> </dd>
      </dl>      </td>
    </tr>
    <tr>
      <td colspan="3" height="21"><span class="style2">* </span><%= notice7 %></td>
    </tr>
    <tr>
      <td colspan="3" height="21">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" height="21">&nbsp;</td>
    </tr>
  </table>
</form>