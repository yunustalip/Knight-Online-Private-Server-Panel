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
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<!--#include file="kontrol.asp"-->
<%
sqlquery = "SELECT * FROM gop_iletisim where id=" & request.querystring("id") & ""
set rs = server.createobject("ADODB.Recordset")
rs.open sqlquery , baglanti , 1 , 3
%>
<title>Mesaj Oku</title>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9">

<style type="text/css">
<!--
.style6 {
	font-size: 10px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
  <tr> 
    <td valign="top" bgcolor="ececec"> 
    <div align="center" class="style6"><strong>Mesaj</strong></div></td>
  </tr>
  <tr valign="top"> 
    <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="style5">&nbsp;</td>
        </tr>
        <tr> 
          <td class="style5"> 

            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><div align="center"> 
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr> 
                        <td width="70%"> <table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr> 
                              <td><span class="style6"><strong><b>Yollayan:</b></strong> 
                                <%= rs("adi") %><br />
<strong><b>Mail:</b></strong> 
                                <%= rs("mail") %><br />
<strong><b>Telefon:</b></strong> 
                                <%= rs("tel") %><br />
<strong><b>Yaþ:</b></strong> 
                                <%= rs("yas") %><br />
<strong><b>Tarih:</b></strong> 
                                <%= rs("tarih") %>
                              
                                <br><br>
                                <%=rs("mesaj")%><br>                              
                                </span></td>
                            </tr>
                            <tr> 
                              <td><div align="center"></div></td>
                            </tr>
                        </table></td>
                      </tr>
                    </table>
                    </div></td>
              </tr>
            </table>
            <%
rs.close 
set rs=nothing 
%>          </td>
        </tr>
      </table>
   </td>
  </tr>
</table>
<!--#include file="alt.asp"-->