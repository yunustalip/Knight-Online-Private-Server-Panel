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
<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "altkategori_sil" then
call altkategori_sil
elseif islem = "altkategori_duzenle" then
call altkategori_duzenle
elseif islem = "altkategori_ekle" then
call altkategori_ekle
elseif islem = "yeni" then
call yeni
elseif islem = "duzenle" then
call duzenle
elseif islem = "" then
call default
end if
sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/altkategori.png" width="128" height="128" align="middle" /><span class="style6"> Alt Kategori Y�netimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Alt Kategori Ad�</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Tan�m</strong></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>ID</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Kategori</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>D�zenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_kat order by katadi asc limit 0,999"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Alt Kategori Yok"
else
for k=1 to "100"
if rs.eof then exit for
%>
                <tr align="left"  bgcolor="#<%
if k mod 2 then
Response.Write "ffffff"
Else
Response.Write "fbe8a6"
end if %>" onmouseover="this.style='BACKGROUND-COLOR: #e58e4d;';" onmouseout="this.style='BACKGROUND-COLOR: #<%
if k mod 2 then
Response.Write "ffffff"
Else
Response.Write "fbe8a6"
end if %>;';">
                  <td height="25" align="center"><strong><%=k%></strong></td>
                  <td ><%=rs("katadi")%></td>
                  <td><%=rs("katbilgi")%></td>
                  <td align="center"><%=rs("katid")%></td>
<td align="center">                  <%
SQL2 ="SELECT * FROM gop_anakat where ankid=" & rs("ankid")
set rs2 = server.createobject("ADODB.Recordset")
rs2.open SQL2 , Baglanti
if rs2.eof or rs2.bof then
Response.Write "Kategorisi Bulunamad�"
else
response.Write "<b>"& rs2("ankadi")&"</b>"
rs2.close
set rs2 = nothing
end if
%></td>

                  <td align="center"><% Response.Write "<a href=altkategoriler.asp?islem=duzenle&katid="&rs("katid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=altkategoriler.asp?islem=altkategori_sil&katid="&rs("katid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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

sub yeni
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/altkategori.png" width="128" height="128" align="middle" /> <span class="style6">Alt Kategori Y�netimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="altkategoriler.asp?islem=altkategori_ekle">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Yeni Kategori Ekle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="41%" height="25" bgcolor="#FFFFFF"><div align="right"><strong>Alt Kategori Ad� </strong></div></td>
                  <td width="1%" bgcolor="#FFFFFF"><div align="right"><strong>:</strong></div></td>
                  <td width="58%" bgcolor="#FFFFFF"><input name="katadi" type="text" class="inputbox" id="katadi" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Alt Kategori Tan�m</strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><textarea name="katbilgi" cols="45" rows="3" class="inputbox2" id="katbilgi"></textarea></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Kategori</strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="fbe8a6"><select name="ankid" class="inputbox" id="ankid">
                    <%
SQL2 ="SELECT * FROM gop_anakat"
set rs2 = server.createobject("ADODB.Recordset")
rs2.open SQL2 , Baglanti
if rs2.eof or rs2.bof then
response.Write "Kategori Yok"
else
do while not rs2.eof

response.Write " <option value="&rs2("ankid")&">"&rs2("ankadi")&"</option>" 

rs2.movenext
loop
end if
rs2.close
set rs2 = nothing
%>
                  </select></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" bgcolor="fbe8a6"><div align="right"></div></td>
                  <td bgcolor="fbe8a6">&nbsp;</td>
                  <td bgcolor="fbe8a6"><input name="Submit" type="submit" class="button" value="Ekle" /></td>
                </tr>
              </table>
              </form></td>
          </tr>
        </table>
 <%
 end sub
 sub duzenle
 %>
 <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/altkategori.png" width="128" height="128" align="middle" /><span class="style6"> AltKategori Y�netimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_kat where katid=" & request.querystring("katid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="altkategoriler.asp?islem=altkategori_duzenle&katid=<%=rs("katid")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Alt Kategori D�zenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="40%" height="25"><div align="right"><strong>ID</strong></div></td>
                  <td width="2%"><div align="right"><strong>:</strong></div></td>
                  <td width="58%" bgcolor="fbe8a6"><%=rs("ankid")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Alt Kategori Ad� </strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><input name="katadi" type="text" class="inputbox" id="katadi" value="<%=rs("katadi")%>" maxlength="50" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Alt Kategori Tan�m</strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="fbe8a6"><textarea name="katbilgi" cols="45" rows="3" class="inputbox2" id="katbilgi"><%=rs("katbilgi")%></textarea></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Kategori</strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><select name="ankid" class="inputbox" id="ankid">
<%
SQL2 ="SELECT * FROM gop_anakat"
set rs2 = server.createobject("ADODB.Recordset")
rs2.open SQL2 , Baglanti
if rs2.eof or rs2.bof then
response.Write "Kategori Yok"
else
do while not rs2.eof

response.Write " <option value="&rs2("ankid")&">"&rs2("ankadi")&"</option>" 

rs2.movenext
loop
end if
rs2.close
set rs2 = nothing
%>
  </select></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25">&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><input name="Submit" type="submit" class="button" value="D�zenle" /></td>
                </tr>
              </table>
                        </form><%
rs.close
set rs = nothing
%></td>
          </tr>
        </table>
<% end sub

sub altkategori_sil
SQL="Delete From gop_kat where katid=" & request.querystring("katid")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub altkategori_duzenle
ankid = Request.Form("ankid")
katadi = Request.Form("katadi")
katbilgi = Request.Form("katbilgi")

baglanti.Execute("UPDATE gop_kat set ankid='"&ankid&"', katbilgi='"&katbilgi&"', katadi='"&katadi&"' where katid='" & request.querystring("katid") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub altkategori_ekle
ankid = Request.Form("ankid")
katadi = Request.Form("katadi")
katbilgi = Request.Form("katbilgi")


SQL="insert into gop_kat (katadi,katbilgi,ankid) values ('"&katadi&"','"&katbilgi&"','"&ankid&"')"
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
%>
        <!--#include file="admin_b.asp"-->