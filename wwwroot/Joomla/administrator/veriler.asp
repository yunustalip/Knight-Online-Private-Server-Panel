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
if islem = "yorum_noonay" then yorum_onay
if islem = "yorum_onay" then yorum_noonay
if islem = "yorum_sil" then yorum_sil
if islem = "" then default

sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Veri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Baþlýk</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Yazan</strong></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>ID</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Kategori</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
listele = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

SQL ="SELECT * FROM gop_veriler order by vid desc LIMIT "& (listele*Sayfa)-(listele) & "," & listele

Set SQLToplam = baglanti.Execute("select count(vid) from gop_veriler") 
TopKayit = SQLToplam(0)

set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Alt Kategori Yok"
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
                  <td ><%=rs("vbaslik")%></td>
                  <td align="center"><% SQL2 ="SELECT * FROM gop_uyeler where uye_id=" & rs("uye_id")
set rs2 = server.createobject("ADODB.Recordset")
rs2.open SQL2 , Baglanti
if rs2.eof or rs2.bof then
response.Write "Üye Silindi"
else
response.Write rs2("uye_adi")

rs2.close
set rs2 = nothing
end if %></td>
                  <td align="center"><%=rs("vid")%></td>
<td align="center">                  <% SQL3 ="SELECT * FROM gop_kat where katid=" & rs("katid")
set rs3 = server.createobject("ADODB.Recordset")
rs3.open SQL3 , Baglanti
if rs3.eof or rs3.bof then
Response.Write "Kategorisi Bulunamadý"
else
response.Write rs3("katadi")
rs3.close
set rs3 = nothing
end if %></td>

                  <td align="center"><% Response.Write "<a href=veriler.asp?islem=duzenle&vid="&rs("vid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=veriler.asp?islem=sil&vid="&rs("vid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
<!--#include file="sayfa.asp"-->
            </td>
          </tr>
        </table>
<%
end sub

sub ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Veri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="veriler.asp?islem=ekle_bitir">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" height="25" bgcolor="#333333" class="style2">Yeni Veri Ekle</td>
                  </tr>
                <tr>
                  <td height="25" bgcolor="#FFFFFF" class="style2"><table width="100%" border="0" cellpadding="3" cellspacing="3">
                    <tr align="left">
                      <td width="48%" valign="middle"><input name="vbaslik" type="text" class="inputbox2" id="vbaslik" size="50" maxlength="60" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <select name="katid" class="inputbox2" id="katid">
                          <%
SQL ="SELECT * FROM gop_kat"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Kategori Yok"
else
do while not rs.eof

response.Write " <option value="&rs("katid")&">"&rs("katadi")&"</option>" 

rs.movenext
loop
end if
rs.close
set rs = nothing
%>
                        </select>
                        &nbsp;
                        <select name="vgoster" class="inputbox2" id="vgoster">
                          <option value="1">Ana Sayfada Göster</option>
                          <option value="0">Ana Sayfada Gizle</option>
                        </select>
                        &nbsp;&nbsp;<strong>Baþlýk Resmi:</strong>
                        <input name="vresim" type="text" class="inputbox2" id="vresim" />
                        <a href="javascript:window.open('upload.asp?islem=veri_resimgonder','yeni','width=400,height=100,scrollbars=yes,statusbar=yes'); void(0);"> Resim Yükle</a></td>
                    </tr>
                    <tr align="left">
                      <td><textarea name="vicerik" id="vicerik" style="width:100%" rows="15"></textarea></td>
                    </tr>
                    <tr align="left">
                      <td>Etiketler : 
                        <input name="vetiket" type="text" class="inputbox2" id="vetiket" size="50" /> 
                        Örn: JoomlASP, Google</td>
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
vbaslik = guvenlik(Request.Form("vbaslik"))
katid = Request.Form("katid")
vgoster = Request.Form("vgoster")
vresim = guvenlik(Request.Form("vresim"))
vicerik = guvenlik(Request.Form("vicerik"))
uye_id = session("uye_id")
vtarih = tarih2
vetiket = guvenlik(Request.Form("vetiket"))


SQL="insert into gop_veriler (vbaslik,katid,vgoster,vresim,vicerik,uye_id,vtarih,vetiket) values ('"&vbaslik&"','"&katid&"','"&vgoster&"','"&vresim&"','"&vicerik&"','"&uye_id&"','"&vtarih&"', '"&vetiket&"')"
Baglanti.Execute(SQL)
Response.Redirect "veriler.asp"
end sub

sub sil
SQL="Delete From gop_veriler where vid=" & request.querystring("vid")
Baglanti.Execute(SQL)
Response.Redirect "veriler.asp"
end sub

sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" width="128" height="128" align="middle" /><span class="style6"> Veri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>       <td height="20">
            			<%
SQL ="SELECT * FROM gop_veriler where vid=" & request.querystring("vid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>

<form action="veriler.asp?islem=duzenle_bitir&vid=<%=rs("vid")%>" method="post" name="verim">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" height="25" bgcolor="#333333" class="style2">Veri Düzenle</td>
                  </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><table width="100%" border="0" cellpadding="3" cellspacing="3">
                    <tr align="left">
                      <td width="48%" valign="middle"><input name="vbaslik" type="text" class="inputbox2" id="vbaslik" value="<%=rs("vbaslik")%>" size="50" maxlength="60" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <select name="katid" class="inputbox2" id="katid">
                          <%
SQL2 ="SELECT * FROM gop_kat"
set rs2 = server.createobject("ADODB.Recordset")
rs2.open SQL2 , Baglanti
if rs2.eof or rs2.bof then
response.Write "Kategori Yok"
else
do while not rs2.eof

if rs2("katid") = rs("katid") then
Response.Write"<option value="&rs2("katid")&" selected=""selected"">"&rs2("katadi")&"</option>"
else
response.Write " <option value="&rs2("katid")&">"&rs2("katadi")&"</option>" 
end if 

rs2.movenext
loop
end if
rs2.close
set rs2 = nothing
%>
                        </select>
                        &nbsp;
                        <select name="vgoster" class="inputbox2" id="vgoster">
<%
if rs("vgoster") = "1" then
Response.Write "<option value=""1"" selected=""selected"">Göster</option>"
else
Response.Write "<option value=""0"" selected=""selected"">Gösterme</option>"
end if
%>                          <option value="1">Ana Sayfada Göster</option>
                          <option value="0">Ana Sayfada Gizle</option>
                        </select>
                        &nbsp;&nbsp;<strong>Baþlýk Resmi:</strong>
                        <input name="vresim" type="text" class="inputbox2" id="vresim" value="<%=rs("vresim")%>" />
                        <a href="javascript:window.open('upload.asp?islem=veri_resimgonder','yeni','width=400,height=100,scrollbars=yes,statusbar=yes'); void(0);"> Resim Yükle</a></td>
                    </tr>
                    <tr align="left">
                      <td><textarea name="vicerik" id="vicerik" style="width:100%" rows="15"><%=rs("vicerik")%></textarea></td>
                    </tr>
                    <tr align="left">
                      <td>Etiketler :
                        <input name="vetiket" type="text" class="inputbox2" id="vetiket" value="<%=rs("vetiket")%>" size="50" />
Örn: JoomlASP, Google</td>
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
vbaslik = guvenlik(Request.Form("vbaslik"))
katid = Request.Form("katid")
vgoster = Request.Form("vgoster")
vresim = guvenlik(Request.Form("vresim"))
vicerik = guvenlik(Request.Form("vicerik"))
vetiket = guvenlik(Request.Form("vetiket"))

baglanti.Execute("UPDATE gop_veriler set vbaslik='"&vbaslik&"', katid='"&katid&"', vgoster='"&vgoster&"', vresim='"&vresim&"', vicerik='"&vicerik&"', vetiket='"&vetiket&"' where vid='" & request.querystring("vid") & "';")
response.Redirect "veriler.asp"
end sub
%>
<!--#include file="admin_b.asp"-->