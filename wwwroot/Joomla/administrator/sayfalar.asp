<!--#include file="admin_a.asp"-->
<script language="javascript" type="text/javascript" src="jscripts/tiny_mce/tiny_mce.js"></script>
<script language="javascript" type="text/javascript" src="tiny_mces.js"></script>
<%
islem = request.querystring("islem")
if islem = "sil" then sil
if islem = "duzenle" then duzenle
if islem = "ekle" then ekle
if islem = "duzenle_bitir" then duzenle_bitir
if islem = "ekle_bitir" then ekle_bitir
if islem = "" then default

sub default
%>

 <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Sayfa Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Baþlýk</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Hit</strong></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>ID</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Tarih</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_sayfa order by sayfaid desc limit 0,100"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write ""
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
                  <td ><%=rs("sayfa_baslik")%></td>
                  <td align="center"><%=rs("sayfa_hit")%></td>
                  <td align="center"><%=rs("sayfaid")%></td>
<td align="center"><%=rs("sayfa_tarih")%></td>

                  <td align="center"><% Response.Write "<a href=sayfalar.asp?islem=duzenle&sayfaid="&rs("sayfaid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=sayfalar.asp?islem=sil&sayfaid="&rs("sayfaid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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

sub ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Sayfa Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="sayfalar.asp?islem=ekle_bitir">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" height="25" bgcolor="#333333" class="style2">Yeni Sayfa Ekle</td>
                  </tr>
                <tr>
                  <td height="25" bgcolor="#FFFFFF" class="style2"><table width="100%" border="0" cellpadding="3" cellspacing="3">
                    <tr align="left">
                      <td width="48%" valign="middle">Baþlýk : 
                        <input name="sayfa_baslik" type="text" class="inputbox2" id="sayfa_baslik" size="50" maxlength="60" /></td>
                    </tr>
                    <tr align="left">
                      <td><textarea name="sayfa_icerik" id="sayfa_icerik" style="width:100%" rows="15"></textarea></td>
                    </tr>
                    <tr align="left">
                      <td><p>
                        <input name="Submit2" type="submit" class="button" value="Ekle" onclick="tinyMCE.triggerSave();" />
                      </p>                        </td>
                    </tr>
                  </table></td>
                </tr>
              </table>
              </form></td>
          </tr>
        </table>
<%
end sub

sub ekle_bitir
sayfa_baslik = Request.Form("sayfa_baslik")
sayfa_icerik = Request.Form("sayfa_icerik")
sayfa_tarih = tarih2


SQL="insert into gop_sayfa (sayfa_baslik,sayfa_icerik,sayfa_tarih) values ('"&sayfa_baslik&"','"&sayfa_icerik&"','"&sayfa_tarih&"')"
Baglanti.Execute(SQL)
Response.Redirect "sayfalar.asp"
end sub

sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" width="128" height="128" align="middle" /><span class="style6"> Sayfa Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>       <td height="20">
            			<%
SQL ="SELECT * FROM gop_sayfa where sayfaid=" & request.querystring("sayfaid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>

<form action="sayfalar.asp?islem=duzenle_bitir&sayfaid=<%=rs("sayfaid")%>" method="post" name="sayfam">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" height="25" bgcolor="#333333" class="style2">Sayfa Düzenle</td>
                  </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><table width="100%" border="0" cellpadding="3" cellspacing="3">
                    <tr align="left">
                      <td width="48%" valign="middle"><input name="sayfa_baslik" type="text" class="inputbox2" id="sayfa_baslik" value="<%=rs("sayfa_baslik")%>" size="50" maxlength="60" /></td>
                    </tr>
                    <tr align="left">
                      <td><textarea name="sayfa_icerik" id="sayfa_icerik" style="width:100%" rows="15"><%=rs("sayfa_icerik")%></textarea></td>
                    </tr>
                    <tr align="left">
                      <td><input name="Submit2" type="submit" class="button" value="Düzenle" onclick="tinyMCE.triggerSave();" /></td>
                    </tr>
                  </table>                  </td>
                  </tr>
              </table>
              </form>
<%

rs.close
set rs = nothing
%>
              
              </td>
          </tr>
        </table>
<%
end sub

sub duzenle_bitir
sayfa_baslik = Request.Form("sayfa_baslik")
sayfa_icerik = Request.Form("sayfa_icerik")

baglanti.Execute("UPDATE gop_sayfa set sayfa_baslik='"&sayfa_baslik&"', sayfa_icerik='"&sayfa_icerik&"' where sayfaid='" & request.querystring("sayfaid") & "';")
response.Redirect "sayfalar.asp"
end sub

sub sil
SQL="Delete From gop_sayfa where sayfaid=" & request.querystring("sayfaid")
Baglanti.Execute(SQL)
Response.Redirect "sayfalar.asp"
end sub
%>
 <!--#include file="admin_b.asp"-->