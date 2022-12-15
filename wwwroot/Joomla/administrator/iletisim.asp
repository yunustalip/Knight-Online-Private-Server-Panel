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
<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "sil" then
call sil
elseif islem = "" then
call default
end if
sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/iletisim.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Ýletiþim Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Adý</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Mail</strong></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Yaþ</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Tel</strong></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Tarih</strong></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Oku</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_iletisim order by id desc limit 0,100"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Mesaj Yok"
else
for k=1 to "100"
if rs.eof then exit for
%>
<tr align="left"  bgcolor="#<%
if k mod 2 then
Response.Write "ffffff"
Else
Response.Write "fbe8a6"
end if %>" onMouseOver="this.style='BACKGROUND-COLOR: #e58e4d;';" onMouseOut="this.style='BACKGROUND-COLOR: #<%
if k mod 2 then
Response.Write "ffffff"
Else
Response.Write "fbe8a6"
end if %>;';">
                  <td height="25" align="center"><strong><%=k%></strong></td>
                  <td >
                  <%
response.Write rs("adi")
%></td>
                  <td align="center"><%
				  response.Write rs("mail") %></td>
                  <td align="center"><% response.Write rs("yas") %></td>
                  <td align="center">
				  <%
response.Write rs("tel")
%>	</td>
                  <td align="center"><%
response.Write rs("tarih")
%>
                  </td>
                  <td align="center"><a href="#" onClick="MM_openBrWindow('mesaj_oku.asp?id=<%=rs("id")%>','mesajlar','scrollbars=yes,width=400,height=450')">Oku</a></td>
                  <td align="center"><% Response.Write "<a href=iletisim.asp?islem=sil&id="&rs("id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>

            </td>
          </tr>
        </table>
<%
end sub

sub sil
SQL="Delete From gop_iletisim where id=" & request.querystring("id")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
%>
<!--#include file="admin_b.asp"-->