<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "yan_gizle" then  yan_gizle
if islem = "yan_goster" then yan_goster
if islem = "ust_gizle" then ust_gizle
if islem = "ust_goster" then ust_goster
if islem = "sira_yukari" then sira_yukari
if islem = "sira_asagi" then sira_asagi
if islem = "sil" then sil
if islem = "duzenle" then duzenle
if islem = "duzenle_bitir" then duzenle_bitir
if islem = "menu_ekle" then menu_ekle
if islem = "menu_ekle_bitir" then menu_ekle_bitir
if islem = "sayfa_menu_ekle" then sayfa_menu_ekle
if islem = "veri_menu_ekle" then veri_menu_ekle
if islem = "" then default

sub default

%><table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/menu.png" width="128" height="128" align="middle" /><span class="style6"> Menü Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Menü Adý</strong></span></td>
                  <td width="50" align="center" bgcolor="#333333"><span class="style4"><strong>ID</strong></span></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Link</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Sýra</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Üst Menü</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Yan Menü</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_menu order by m_order asc limit 0,999"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Menü Yok"
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
                  <td ><%=rs("m_adi")%></td>
                  <td align="center"><%=rs("m_id")%></td>
                  <td><%=rs("m_link")%></td>
                  <td align="center"><a href="menuler.asp?islem=sira_yukari&m_id=<%=rs("m_id")%>"><img src="../images/yukari.png" border="0"> </a><%=rs("m_order")%> <a href="menuler.asp?islem=sira_asagi&m_id=<%=rs("m_id")%>"><img src="../images/asagi.png" border="0"></a></td>
<%
if rs("m_ust") = "1" then
Response.Write "<td align=""center""><a href=menuler.asp?islem=ust_gizle&m_id="&rs("m_id")&"><img src=""../images/yayinda.png"" border=0></a></td>"
else
Response.Write "<td align=""center""><a href=menuler.asp?islem=ust_goster&m_id="&rs("m_id")&"><img src=""../images/yayinda_degil.png"" border=0></a></td>"
end if
if rs("m_yan") = "1" then
Response.Write "<td align=""center""><a href=menuler.asp?islem=yan_gizle&m_id="&rs("m_id")&"><img src=""../images/yayinda.png"" border=0></a></td>"
else
Response.Write "<td align=""center""><a href=menuler.asp?islem=yan_goster&m_id="&rs("m_id")&"><img src=""../images/yayinda_degil.png"" border=0></a></td>"
end if
%>
                  <td align="center"><% Response.Write "<a href=menuler.asp?islem=duzenle&m_id="&rs("m_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>
                    </td>
                  <td align="center"><%Response.Write "<a href=menuler.asp?islem=sil&m_id="&rs("m_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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

sub yan_gizle
baglanti.Execute("UPDATE gop_menu set m_yan='"& 0 &"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end sub

sub yan_goster
baglanti.Execute("UPDATE gop_menu set m_yan='"& 1 &"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end sub

sub ust_gizle
baglanti.Execute("UPDATE gop_menu set m_ust='"& 0 &"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end sub

sub ust_goster
baglanti.Execute("UPDATE gop_menu set m_ust='"& 1 &"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end sub

sub sira_yukari
Set rs = baglanti.Execute("Select * from gop_menu where m_id=" & request.querystring("m_id") & " ;")
if rs.eof or rs.bof then
Response.Redirect "hata2.asp"
else
baglanti.Execute("UPDATE gop_menu set m_order='"& rs("m_order") - 1 &"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end if
end sub

sub sira_asagi
Set rs = baglanti.Execute("Select * from gop_menu where m_id=" & request.querystring("m_id") & " ;")
if rs.eof or rs.bof then
Response.Redirect "hata2.asp"
else
baglanti.Execute("UPDATE gop_menu set m_order='"& rs("m_order") + 1 &"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end if
end sub

sub sil
SQL="Delete From gop_menu where m_id=" & request.querystring("m_id")
Baglanti.Execute(SQL)
Response.Redirect "menuler.asp"
end sub

sub duzenle_bitir
m_adi = Request.Form("m_adi")
m_link = Request.Form("m_link")
m_order = Request.Form("m_order")
m_yan = Request.Form("m_yan")
m_ust = Request.Form("m_ust")

baglanti.Execute("UPDATE gop_menu set m_adi='"&m_adi&"', m_link='"&m_link&"',m_order='"&m_order&"', m_yan='"&m_yan&"', m_ust='"&m_ust&"' where m_id='" & request.querystring("m_id") & "';")
response.Redirect "menuler.asp"
end sub

sub menu_ekle_bitir
m_adi = Request.Form("m_adi")
m_link = Request.Form("m_link")
m_order = Request.Form("m_order")
m_yan = Request.Form("m_yan")
m_ust = Request.Form("m_ust")
dlink = Request.Form("dlink")

if m_link = "" then
SQL="insert into gop_menu (m_adi,m_link,m_order,m_yan,m_ust) values ('"&m_adi&"','"&dlink&"','"&m_order&"','"&m_yan&"','"&m_ust&"')"
Baglanti.Execute(SQL)
Response.Redirect "menuler.asp"
else
SQL="insert into gop_menu (m_adi,m_link,m_order,m_yan,m_ust) values ('"&m_adi&"','"&m_link&"','"&m_order&"','"&m_yan&"','"&m_ust&"')"
Baglanti.Execute(SQL)
Response.Redirect "menuler.asp"
end if
end sub

sub sayfa_menu_ekle
m_adi = Request.Form("m_adi")
m_order = Request.Form("m_order")
m_yan = Request.Form("m_yan")
m_ust = Request.Form("m_ust")
slink = Request.Form("slink")


SQL="insert into gop_menu (m_adi,m_link,m_order,m_yan,m_ust) values ('"&m_adi&"','"&slink&"','"&m_order&"','"&m_yan&"','"&m_ust&"')"
Baglanti.Execute(SQL)
Response.Redirect "menuler.asp"
end sub

sub veri_menu_ekle
m_adi = Request.Form("m_adi")
m_order = Request.Form("m_order")
m_yan = Request.Form("m_yan")
m_ust = Request.Form("m_ust")
vlink = Request.Form("vlink")


SQL="insert into gop_menu (m_adi,m_link,m_order,m_yan,m_ust) values ('"&m_adi&"','"&vlink&"','"&m_order&"','"&m_yan&"','"&m_ust&"')"
Baglanti.Execute(SQL)
Response.Redirect "menuler.asp"
end sub

sub menu_ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/menu.png" width="128" height="128" align="middle" /> <span class="style6">Menü Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20">
      
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="33%" height="25" class="baslik2"><div align="center" class="current" id="t11"><a href="#" onclick="showtab(1,1); return false;">Normal</a></div></td>
    <td width="33%" class="baslik2"><div align="center" id="t12"><a href="#" onclick="showtab(1,2); return false;">Sayfadan</a></div></td>
    <td width="33%" class="baslik2"><div align="center" id="t13"><a href="#" onclick="showtab(1,3); return false;">Veriden</a></div></td>
  </tr>
</table>
<div id="tab11">
      <form id="form1" name="form1" method="post" action="menuler.asp?islem=menu_ekle_bitir">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Yeni Menü Ekle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="30%" height="25"><div align="right"><strong>Menü Adý </strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_adi" type="text" class="inputbox" id="m_adi" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Link</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><input name="m_link" type="text" class="inputbox" id="m_link" size="50" /> 
                    Dahili link kullanacaksanýz boþ býrakýnýz.</td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" bgcolor="#FF9999"><div align="right"><strong>Dahili Link</strong></div></td>
                  <td bgcolor="#FF9999"><strong>:</strong></td>
                  <td bgcolor="#FF9999"><select name="dlink" id="dlink">
<option selected="selected">Boþ</option>
<% 
SQL ="SELECT * FROM gop_eklentiler"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Bileþen Bulunamadý"
else
do while not rs.eof

response.Write " <option value=""default.asp?islem=bilesen&component="&rs("eklenti_k")&""">"&rs("eklenti_adi")&"</option>" 

rs.movenext
loop
end if
rs.close
set rs = nothing
%>
                  </select>
                  </td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Üst Menü</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><select name="m_ust" class="inputbox" id="m_ust">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select>                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Yan Menü</strong></div></td>
                  <td>&nbsp;</td>
                  <td bgcolor="#FFFFFF"><select name="m_yan" class="inputbox" id="m_yan">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Sýra</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_order" type="text" class="inputbox" id="m_order" size="5" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"></div></td>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td bgcolor="#FFFFFF"><input name="Submit" type="submit" class="button" value="Ekle" /></td>
                </tr>
              </table>
              </form>
              </div>
              <div id="tab12" style="display:none">
              <form id="form2" name="form2" method="post" action="menuler.asp?islem=sayfa_menu_ekle">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Sayfadan Menü Oluþtur</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="30%" height="25"><div align="right"><strong>Menü Adý </strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_adi" type="text" class="inputbox" id="m_adi" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Sayfa</strong></div></td>
                  <td><strong>:</strong></td>
                  <td>
<select name="slink" id="slink">

<% 
SQL ="SELECT * FROM gop_sayfa"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Sayfa Bulunamadý"
else
do while not rs.eof

response.Write " <option value=""default.asp?islem=bilesen&component=sayfa_sistemi&sayfaid="&rs("sayfaid")&"&sayfa_adi="&rs("sayfa_baslik")&""">"&rs("sayfa_baslik")&"</option>" 

rs.movenext
loop
end if
rs.close
set rs = nothing
%>
                  </select>                  </td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Üst Menü</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><select name="m_ust" class="inputbox" id="m_ust">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select>                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Yan Menü</strong></div></td>
                  <td>&nbsp;</td>
                  <td bgcolor="#FFFFFF"><select name="m_yan" class="inputbox" id="m_yan">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Sýra</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_order" type="text" class="inputbox" id="m_order" size="5" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"></div></td>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td bgcolor="#FFFFFF"><input name="Submit" type="submit" class="button" value="Ekle" /></td>
                </tr>
              </table>
              </form>
              </div>
              <div id="tab13" style="display:none">
              <form id="form3" name="form3" method="post" action="menuler.asp?islem=veri_menu_ekle">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Veriden Menü Oluþtur</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="30%" height="25"><div align="right"><strong>Menü Adý </strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_adi" type="text" class="inputbox" id="m_adi" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Veri</strong></div></td>
                  <td><strong>:</strong></td>
                  <td>
<select name="vlink" id="vlink">
<% 
SQL ="SELECT * FROM gop_veriler"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Veri Bulunamadý"
else
do while not rs.eof

response.Write " <option value=""default.asp?islem=oku&vid="&rs("vid")&"&baslik="&rs("vbaslik")&""">"&rs("vbaslik")&"</option>" 

rs.movenext
loop
end if
rs.close
set rs = nothing
%>
                  </select>                  </td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Üst Menü</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><select name="m_ust" class="inputbox" id="m_ust">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select>                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Yan Menü</strong></div></td>
                  <td>&nbsp;</td>
                  <td bgcolor="#FFFFFF"><select name="m_yan" class="inputbox" id="m_yan">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Sýra</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_order" type="text" class="inputbox" id="m_order" size="5" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"></div></td>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td bgcolor="#FFFFFF"><input name="Submit" type="submit" class="button" value="Ekle" /></td>
                </tr>
              </table>
              </form>
              </div></td>
          </tr>
        </table>
<%
end sub

sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/menu.png" width="128" height="128" align="middle" /><span class="style6">Menü Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_menu where m_id=" & request.querystring("m_id")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="menuler.asp?islem=menu_duzenle&m_id=<%=rs("m_id")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Menü Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25"><div align="right"><strong>Menu ID</strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("m_id")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Menü Adý </strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><input name="m_adi" type="text" class="inputbox" id="m_adi" value="<%=rs("m_adi")%>" maxlength="50" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Link</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="m_link" type="text" class="inputbox" id="m_link" value="<%=rs("m_link")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Yan Menü</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><select name="m_yan" class="inputbox" id="m_yan">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Üst Menü</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><select name="m_ust" class="inputbox" id="m_ust">
                    <option value="1">Göster</option>
                    <option value="0">Gösterme</option>
                  </select></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Sýra</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><input name="m_order" type="text" class="inputbox" id="m_order" value="<%=rs("m_order")%>" size="2" maxlength="2" /></td>
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