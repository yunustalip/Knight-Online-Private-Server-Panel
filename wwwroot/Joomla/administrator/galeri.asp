<!--#include file="admin_a.asp"-->
<% 
listele = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

islem = request.querystring("islem")
if islem = "sil" then sil
if islem = "duzenle" then duzenle
if islem = "duzenle_bitir" then duzenle_bitir
if islem = "ekle" then ekle
if islem = "resimler" then resimler
if islem = "resim_ekle" then resim_ekle
if islem = "resim_sil" then resim_sil
if islem = "resim_duzenle" then resim_duzenle
if islem = "resim_duzenle_bitir" then resim_duzenle_bitir
if islem = "" then default

sub default
 %>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/galeri.png" width="128" height="128" align="middle" /><span class="style6"> Galeri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20" bgcolor="e58e4d"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="50" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Kategoriler</strong></span><span class="style4"><strong></strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%

set rs = baglanti.Execute("select * from gop_galerikat order by galkat asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)
Set SQLToplam = baglanti.Execute("select count(galid) from gop_galerikat") 
TopKayit = SQLToplam(0)
if rs.eof or rs.bof then
response.Write "Kategori Bulunamadý"
else
for k=1 to "50"
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
                  <td align="center"><strong><%=k%></strong></td>
                  <td height="20" >
<strong><%=rs("galkat")%>
                    <br /></td>
                  <td align="center"><% Response.Write "<a href=galeri.asp?islem=duzenle&galid="&rs("galid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=galeri.asp?islem=sil&galid="&rs("galid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
              <!--#include file="sayfa.asp"--></td>
          </tr>
          <tr bgcolor="ffffff">
            <td height="25"><form id="form1" name="form1" method="post" action="galeri.asp?islem=ekle">
              &nbsp;
              <input name="galkat" type="text" class="inputbox2" id="galkat" />
              <input name="button" type="submit" class="button" id="button" value="Kategori Ekle" />
            </form>
            </td>
          </tr>
        </table>
 <%
end sub

sub ekle
galkat = Request.Form("galkat")

SQL="insert into gop_galerikat (galkat) values ('"&galkat&"')"
Baglanti.Execute(SQL)
Response.Redirect "galeri.asp"
end sub

sub sil
SQL="Delete From gop_galerikat where galid=" & request.querystring("galid")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub duzenle_bitir
galkat = Request.Form("galkat")
baglanti.Execute("UPDATE gop_galerikat set galkat='"&galkat&"' where galid='" & request.querystring("galid") & "';")
Response.Redirect "galeri.asp"
end sub

sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/galeri.png" width="128" height="128" align="middle" /><span class="style6"> Galeri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_galerikat where galid=" & request.querystring("galid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="galeri.asp?islem=duzenle_bitir&galid=<%=rs("galid")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2"> Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25" class="style3"><div align="right"><strong> ID</strong></div></td>
                  <td width="1%" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("galid")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" class="style3"><div align="right"><strong>Kategori Adý </strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><input name="galkat" type="text" class="inputbox2" id="galkat" value="<%=rs("galkat")%>" size="60" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"></div></td>
                  <td bgcolor="fbe8a6">&nbsp;</td>
                  <td bgcolor="fbe8a6"><input name="Submit" type="submit" class="button" value="Düzenle" /></td>
                </tr>
              </table>
                        </form><%
rs.close
set rs = nothing
%></td>
          </tr>
        </table>
<%
end sub

sub resimler
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/galeri.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Galeri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20" bgcolor="e58e4d"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Resim Adý</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Resim Adres</strong></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Hit</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Kategori</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_galeri order by resid desc LIMIT "& (listele*Sayfa)-(listele) & "," & listele
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Resim Yok"
else
for k=1 to "50"
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
                  <td ><%=rs("resadi")%></td>
                  <td align="center"><%=rs("resresim")%></td>
                  <td align="center"><%=rs("rhit")%></td>
<td align="center">                  <% SQL2 ="SELECT * FROM gop_galerikat where galid=" & rs("galid")
set rs2 = server.createobject("ADODB.Recordset")
rs2.open SQL2 , Baglanti
if rs2.eof or rs2.bof then
response.Write "Kategori Silindi"
else
response.Write rs2("galkat")

rs2.close
set rs2 = nothing
end if %></td>

                  <td align="center"><% Response.Write "<a href=galeri.asp?islem=resim_duzenle&resid="&rs("resid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=galeri.asp?islem=resim_sil&resid="&rs("resid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
              <!--#include file="sayfa.asp"-->            </td>
          </tr>
          <tr>
            <td height="20" bgcolor="#CC0000" class="style2">&nbsp;Resim Ekle</td>
          </tr>
          <tr>
            <td height="30" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="galeri.asp?islem=resim_ekle">
&nbsp;Kategori:
<select name="galid" class="inputbox2" id="galid">
                          <%
SQL ="SELECT * FROM gop_galerikat"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Kategori Yok"
else
do while not rs.eof

response.Write " <option value="&rs("galid")&">"&rs("galkat")&"</option>" 

rs.movenext
loop
end if
rs.close
set rs = nothing
%>
                    </select>
Resim Adý:
<input name="resadi" type="text" class="inputbox2" id="resadi" />            
                                    Resim Adres: 
                                    <input name="resresim" type="text" class="inputbox2" id="resresim" />
             <a href="javascript:window.open('upload.asp?islem=gonder','yeni','width=400,height=100,scrollbars=yes,statusbar=yes'); void(0);"> Resim Yükle</a> 
             <input name="button" type="submit" class="button" id="button" value="Kaydet" />
             </form>
            </td>
          </tr>
        </table>
<%
end sub

sub resim_ekle
galid = Request.Form("galid")
resadi = Request.Form("resadi")
resresim = Request.Form("resresim")


SQL="insert into gop_galeri (galid,resadi,resresim) values ('"&galid&"','"&resadi&"','"&resresim&"')"
Baglanti.Execute(SQL)
Response.Redirect "galeri.asp?islem=resimler"
end sub

sub resim_sil
SQL="Delete From gop_galeri where resid=" & request.querystring("resid")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub resim_duzenle_bitir
galid = Request.Form("galid")
resadi = Request.Form("resadi")
resresim = Request.Form("resresim")
baglanti.Execute("UPDATE gop_galeri set galid='"&galid&"', resadi='"&resadi&"', resresim='"&resresim&"' where resid='" & request.querystring("resid") & "';")
Response.Redirect "galeri.asp?islem=resimler"
end sub

sub resim_duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/galeri.png" width="128" height="128" align="middle" /><span class="style6"> Galeri  Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_galeri where resid=" & request.querystring("resid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="galeri.asp?islem=resim_duzenle_bitir&resid=<%=rs("resid")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2"> Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25" class="style3"><div align="right"><strong> ID</strong></div></td>
                  <td width="1%" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("resid")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" class="style3"><div align="right"><strong>Resim Adý </strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><input name="resadi" type="text" class="inputbox2" id="resadi" value="<%=rs("resadi")%>" size="60" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25" class="style3"><div align="right"><strong>Resim Adres</strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="fbe8a6"><input name="resresim" type="text" class="inputbox2" id="resresim" value="<%=rs("resresim")%>" size="60" /> 
                  Yükle</td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25" bgcolor="ffffff" class="style3"><div align="right"><strong>Kategori</strong></div></td>
                  <td bgcolor="ffffff" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="ffffff"><select name="galid" class="inputbox2" id="galid">
<% SQL3 ="SELECT * FROM gop_galerikat"
set rs3 = server.createobject("ADODB.Recordset")
rs3.open SQL3 , Baglanti
if rs3.eof or rs3.bof then
Response.Write "Kategorisi Bulunamadý"
else
do while not rs3.eof
response.Write " <option value="&rs3("galid")&">"&rs3("galkat")&"</option>" 
rs3.movenext
loop
end if
rs3.close
set rs3 = nothing %></select></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"></div></td>
                  <td bgcolor="fbe8a6">&nbsp;</td>
                  <td bgcolor="fbe8a6"><input name="Submit" type="submit" class="button" value="Düzenle" /></td>
                </tr>
              </table>
                        </form><%
rs.close
set rs = nothing
%></td>
          </tr>
        </table>
<%
end sub
%>
        <!--#include file="admin_b.asp"-->
