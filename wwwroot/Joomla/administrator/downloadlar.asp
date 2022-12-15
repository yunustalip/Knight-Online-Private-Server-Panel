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
if islem = "ekle" then
call ekle
elseif islem = "ekle_ok" then
call ekle_ok
elseif islem = "onay" then
call onay
elseif islem = "duzenle" then
call duzenle
elseif islem = "duzenle_ok" then
call duzenle_ok
elseif islem = "kategoriler" then
call kategoriler
elseif islem = "kategori_ekle" then
call kategori_ekle
elseif islem = "kategori_duzenle" then
call kategori_duzenle
elseif islem = "kategori_duzenle_ok" then
call kategori_duzenle_ok
elseif islem = "sil" then
call sil
elseif islem = "kategori_sil" then
call kategori_sil
elseif islem = "gizle" then
call gizle
elseif islem = "goster" then
call goster
elseif islem = "" then
call default
end if
sub default
%><table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/downloadlar.png" width="128" height="128" align="middle" /><span class="style6"> Download Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="50" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Download Bilgileri</strong></span><span class="style4"><strong></strong></span></td>
                  <td width="75" align="center" bgcolor="#333333" class="style4"><strong>Kategori</strong></td>
                  <td width="75" align="center" bgcolor="#333333"><span class="style4"><strong>Onay</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Gösterim</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
listele = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

set rs = baglanti.Execute("select * from gop_download where down_onay = '"& 1 &"' order by down_id asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)
Set SQLToplam = baglanti.Execute("select count(link_id) from gop_linkler") 
TopKayit = SQLToplam(0)
if rs.eof or rs.bof then
response.Write "Download Bulunamadý"
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
                  <td ><br />
<a href="<%=rs("down_link")%>" target="_blank"><strong><%=rs("down_adi")%></strong></a><br />
                    <%=rs("down_bilgi")%><br />
                    <br /></td>
                  <td align="center"><% SQL3 ="SELECT * FROM gop_download_kat where dkid=" & rs("dkid")
set rs3 = server.createobject("ADODB.Recordset")
rs3.open SQL3 , Baglanti
if rs3.eof or rs3.bof then
Response.Write "Kategorisi Bulunamadý"
else
response.Write rs3("dk_adi")
rs3.close
set rs3 = nothing
end if %></td>
                  <td align="center"><%
if rs("down_onay") = "1" then
Response.Write "<a href=downloadlar.asp?islem=gizle&down_id="&rs("down_id")&"><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=downloadlar.asp?islem=goster&down_id="&rs("down_id")&"><img src=""../images/yayinda_degil.png"" border=0></a>"
end if%></td>
                  <td align="center"><%=rs("down_hit")%></td>
                                 
                  <td align="center"><% Response.Write "<a href=downloadlar.asp?islem=duzenle&down_id="&rs("down_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=downloadlar.asp?islem=sil&down_id="&rs("down_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
%>
              </table>
              <div align="center">
                <%

If CInt(TopKayit) > CInt(listele) Then 
SayfaSayisi = CInt(TopKayit) / CInt(listele) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a href=?s=1><< Ýlk</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a href=?s="&Sayfa-1&">< Önceki</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a href=?s=" & t & "><b>" & t & "</b></a> "
end if
next 
For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s=" & d & "><b>" & d & "</b></a> " 
end if
Next

if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s="&Sayfa+1&"> Sonraki ></a> "
Response.Write " <a href=?s="&SayfaSayisi&">Son >></a>"
end if
end if
end if
end if%>            
              </div></td>
          </tr>
        </table>
        <% end sub
		sub ekle
		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/downloadlar.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Download Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="downloadlar.asp?islem=ekle_ok">
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="25" colspan="3" bgcolor="#333333" class="style2"> &nbsp;Ekle</td>
              </tr>

              <tr bgcolor="#FFFFFF">
                <td width="42%" height="25" class="style3"><div align="right"><strong>Baþlýk </strong></div></td>
                <td width="1%" class="style3"><div align="right"><strong>:</strong></div></td>
                <td width="57%" bgcolor="#FFFFFF"><input name="down_adi" type="text" class="inputbox2" id="down_adi" size="60" /></td>
              </tr>
              <tr bgcolor="fbe8a6">
                <td height="25" class="style3"><div align="right"><strong>Açýklama</strong></div></td>
                <td class="style3"><div align="right"><strong>:</strong></div></td>
                <td bgcolor="fbe8a6"><textarea name="down_bilgi" cols="60" rows="3" class="inputbox2" id="down_bilgi"></textarea></td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td height="25" class="style3"><div align="right"><strong>URL</strong></div></td>
                <td class="style3"><div align="right"><strong>:</strong></div></td>
                <td bgcolor="#FFFFFF"><input name="down_link" type="text" class="inputbox2" id="down_link" size="60" /></td>
              </tr>
              <tr bgcolor="fbe8a6">
                <td height="25" class="style3"><div align="right"><strong>Resim</strong></div></td>
                <td bgcolor="fbe8a6" class="style3"><div align="right"><strong>:</strong></div></td>
                <td bgcolor="fbe8a6"><input name="down_resim" type="text" class="inputbox2" id="down_resim" />
                  <a href="javascript:window.open('resim_upload.asp','yeni','width=400,height=100,scrollbars=yes,statusbar=yes'); void(0);"> Resim Yükle</a></td>
              </tr>
              <tr bgcolor="fbe8a6">
                <td height="25" bgcolor="ffffff" class="style3"><div align="right"><strong>Kategori</strong></div></td>
                <td bgcolor="ffffff" class="style3"><div align="right"><strong>:</strong></div></td>
                <td bgcolor="ffffff"><select name="dkid" class="inputbox2" id="dkid">
                    <% SQL3 ="SELECT * FROM gop_download_kat"
set rs3 = server.createobject("ADODB.Recordset")
rs3.open SQL3 , Baglanti
if rs3.eof or rs3.bof then
Response.Write "Kategorisi Bulunamadý"
else
do while not rs3.eof
response.Write " <option value="&rs3("dkid")&">"&rs3("dk_adi")&"</option>" 
rs3.movenext
loop
end if
rs3.close
set rs3 = nothing %>
                </select></td>
              </tr>
              <tr bgcolor="fbe8a6">
                <td height="25"><div align="right"></div></td>
                <td bgcolor="fbe8a6">&nbsp;</td>
                <td bgcolor="fbe8a6"><input name="Submit" type="submit" class="button" value="Ekle" /></td>
              </tr>
            </table>
          </form></td>
          </tr>
        </table>
        <%
		end sub
		sub ekle_ok
		down_adi = Request.Form("down_adi")
down_link = Request.Form("down_link")
down_bilgi = guvenlik(Request.Form("down_bilgi"))
down_resim = Request.Form("down_resim")
dkid = Request.Form("dkid")


SQL="insert into gop_download (down_adi,down_link,down_bilgi,down_resim,dkid) values ('"&down_adi&"','"&down_link&"','"&down_bilgi&"','"&down_resim&"','"&dkid&"')"
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
		end sub
		sub duzenle
		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/downloadlar.png" width="128" height="128" align="middle" /><span class="style6">Download Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_download where down_id=" & request.querystring("down_id")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="downloadlar.asp?islem=duzenle_ok&down_id=<%=rs("down_id")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2"> Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25" class="style3"><div align="right"><strong> ID</strong></div></td>
                  <td width="1%" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("down_id")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" class="style3"><div align="right"><strong>Baþlýk </strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><input name="down_adi" type="text" class="inputbox2" id="down_adi" value="<%=rs("down_adi")%>" size="60" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25" class="style3"><div align="right"><strong>Açýklama</strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="fbe8a6"><textarea name="down_bilgi" cols="60" rows="3" class="inputbox2" id="down_bilgi"><%=rs("down_bilgi")%></textarea></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" class="style3"><div align="right"><strong>URL</strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><input name="down_link" type="text" class="inputbox2" id="down_link" value="<%=rs("down_link")%>" size="60" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25" class="style3"><div align="right"><strong>Resim</strong></div></td>
                  <td bgcolor="fbe8a6" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="fbe8a6"><input name="down_resim" type="text" class="inputbox2" id="down_resim" value="<%=rs("down_resim")%>" />
                  <a href="javascript:window.open('resim_upload.asp','yeni','width=400,height=100,scrollbars=yes,statusbar=yes'); void(0);"> Resim Yükle</a></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25" bgcolor="ffffff" class="style3"><div align="right"><strong>Kategori</strong></div></td>
                  <td bgcolor="ffffff" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="ffffff"><select name="dkid" class="inputbox2" id="dkid">
<% SQL3 ="SELECT * FROM gop_download_kat"
set rs3 = server.createobject("ADODB.Recordset")
rs3.open SQL3 , Baglanti
if rs3.eof or rs3.bof then
Response.Write "Kategorisi Bulunamadý"
else
do while not rs3.eof
response.Write " <option value="&rs3("dkid")&">"&rs3("dk_adi")&"</option>" 
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
		sub duzenle_ok
		down_adi = Request.Form("down_adi")
down_link = Request.Form("down_link")
down_bilgi = guvenlik(Request.Form("down_bilgi"))
down_resim = Request.Form("down_resim")
dkid = Request.Form("dkid")

baglanti.Execute("UPDATE gop_download set down_adi='"&down_adi&"', down_link='"&down_link&"',down_bilgi='"&down_bilgi&"',down_resim='"&down_resim&"',dkid='"&dkid&"' where down_id='" & request.querystring("down_id") & "';")
Response.Redirect "downloadlar.asp"
		end sub
		sub kategoriler
		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/downloadlar.png" width="128" height="128" align="middle" /><span class="style6"> Download Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="50" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Kategoriler</strong></span><span class="style4"><strong></strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
listele = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

set rs = baglanti.Execute("select * from gop_download_kat order by dk_adi asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)
Set SQLToplam = baglanti.Execute("select count(link_id) from gop_linkler") 
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
<strong><%=rs("dk_adi")%>
                    <br /></td>
                  <td align="center"><% Response.Write "<a href=downloadlar.asp?islem=kategori_duzenle&dkid="&rs("dkid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=downloadlar.asp?islem=kategori_sil&dkid="&rs("dkid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
%>
              </table>
              <div align="center">
                <%

If CInt(TopKayit) > CInt(listele) Then 
SayfaSayisi = CInt(TopKayit) / CInt(listele) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a href=?s=1><< Ýlk</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a href=?s="&Sayfa-1&">< Önceki</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a href=?s=" & t & "><b>" & t & "</b></a> "
end if
next 
For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s=" & d & "><b>" & d & "</b></a> " 
end if
Next

if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s="&Sayfa+1&"> Sonraki ></a> "
Response.Write " <a href=?s="&SayfaSayisi&">Son >></a>"
end if
end if
end if
end if%>            
              </div></td>
          </tr>
          <tr bgcolor="ffffff">
            <td height="25"><form id="form1" name="form1" method="post" action="downloadlar.asp?islem=kategori_ekle">
              &nbsp;
              <input name="dk_adi" type="text" class="inputbox2" id="dk_adi" />
              <input name="button" type="submit" class="button" id="button" value="Kategori Ekle" />
            </form>
            </td>
          </tr>
        </table>
        <%
		end sub
		sub kategori_ekle
		dk_adi = Request.Form("dk_adi")
SQL="insert into gop_download_kat (dk_adi) values ('"&dk_adi&"')"
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
sub kategori_duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/downloadlar.png" width="128" height="128" align="middle" /><span class="style6">Download Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_download_kat where dkid=" & request.querystring("dkid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="downloadlar.asp?islem=kategori_duzenle_ok&dkid=<%=rs("dkid")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2"> Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25" class="style3"><div align="right"><strong> ID</strong></div></td>
                  <td width="1%" class="style3"><div align="right"><strong>:</strong></div></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("dkid")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25" class="style3"><div align="right"><strong>Kategori Adý </strong></div></td>
                  <td class="style3"><div align="right"><strong>:</strong></div></td>
                  <td bgcolor="#FFFFFF"><input name="dk_adi" type="text" class="inputbox2" id="dk_adi" value="<%=rs("dk_adi")%>" size="60" /></td>
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
		sub kategori_duzenle_ok
		dk_adi = Request.Form("dk_adi")

baglanti.Execute("UPDATE gop_download_kat set dk_adi='"&dk_adi&"' where dkid='" & request.querystring("dkid") & "';")
Response.Redirect "downloadlar.asp?islem=kategoriler"
end sub
sub sil
SQL="Delete From gop_download where down_id=" & request.querystring("down_id")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
sub kategori_sil
SQL="Delete From gop_download_kat where dkid=" & request.querystring("dkid")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
sub gizle
baglanti.Execute("UPDATE gop_download set down_onay='"& 0 &"' where down_id='" & request.querystring("down_id") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub goster
baglanti.Execute("UPDATE gop_download set down_onay='"& 1 &"' where down_id='" & request.querystring("down_id") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
sub onay
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/downloadlar.png" width="128" height="128" align="middle" /><span class="style6"> Download Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="50" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Download Bilgileri</strong></span><span class="style4"><strong></strong></span></td>
                  <td width="75" align="center" bgcolor="#333333" class="style4"><strong>Kategori</strong></td>
                  <td width="75" align="center" bgcolor="#333333"><span class="style4"><strong>Onay</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Gösterim</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
listele = 50
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

set rs = baglanti.Execute("select * from gop_download where down_onay = '"& 0 &"' order by down_id asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)
Set SQLToplam = baglanti.Execute("select count(link_id) from gop_linkler") 
TopKayit = SQLToplam(0)
if rs.eof or rs.bof then
response.Write "Download Bulunamadý"
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
                  <td ><br />
<a href="<%=rs("down_link")%>" target="_blank"><strong><%=rs("down_adi")%></strong></a><br />
                    <%=rs("down_bilgi")%><br />
                    <br /></td>
                  <td align="center"><% SQL3 ="SELECT * FROM gop_download_kat where dkid=" & rs("dkid")
set rs3 = server.createobject("ADODB.Recordset")
rs3.open SQL3 , Baglanti
if rs3.eof or rs3.bof then
Response.Write "Kategorisi Bulunamadý"
else
response.Write rs3("dk_adi")
rs3.close
set rs3 = nothing
end if %></td>
                  <td align="center"><%
if rs("down_onay") = "1" then
Response.Write "<a href=downloadlar.asp?islem=gizle&down_id="&rs("down_id")&"><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=downloadlar.asp?islem=goster&down_id="&rs("down_id")&"><img src=""../images/yayinda_degil.png"" border=0></a>"
end if%></td>
                  <td align="center"><%=rs("down_hit")%></td>
                                 
                  <td align="center"><% Response.Write "<a href=downloadlar.asp?islem=duzenle&down_id="&rs("down_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=downloadlar.asp?islem=sil&down_id="&rs("down_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
              <div align="center">
                <%

If CInt(TopKayit) > CInt(listele) Then 
SayfaSayisi = CInt(TopKayit) / CInt(listele) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a href=?s=1><< Ýlk</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a href=?s="&Sayfa-1&">< Önceki</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a href=?s=" & t & "><b>" & t & "</b></a> "
end if
next 
For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s=" & d & "><b>" & d & "</b></a> " 
end if
Next

if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a href=?s="&Sayfa+1&"> Sonraki ></a> "
Response.Write " <a href=?s="&SayfaSayisi&">Son >></a>"

end if
end if
end if%>            
              </div></td>
          </tr>
        </table>
        <%
		end sub
		%>
<!--#include file="admin_b.asp"-->