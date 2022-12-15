<!--#include file="admin_a.asp"-->

<%
act = request.querystring("act")
if act = "gizle" then
call gizle
elseif act = "goster" then
call goster
elseif act = "sil" then
call sil
elseif act = "duzenle" then
call duzenle
elseif act = "duzenle_bitir" then
call duzenle_bitir
elseif act = "ekle" then
call ekle
elseif act = "ekle_bitir" then
call ekle_bitir
elseif act = "onay" then
call onay
elseif act = "" then
call default
end if

sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/linkler.png" width="128" height="128" align="middle" /><span class="style6"> Link Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="50" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Bilgi</strong></span><span class="style4"><strong></strong></span></td>
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

set rs = baglanti.Execute("select * from gop_linkler where link_onay = '"& 1 &"' order by link_id asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)
Set SQLToplam = baglanti.Execute("select count(link_id) from gop_linkler") 
TopKayit = SQLToplam(0)
if rs.eof or rs.bof then
response.Write "Link Bulunamadý"
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
<a href="http://<%=rs("link_url")%>" target="_blank"><strong><%=rs("link_adi")%></strong></a><br />
                    <%=rs("link_aciklama")%><br />
                    <br />
</td>
                  <td align="center"><%
if rs("link_onay") = "1" then
Response.Write "<a href=linkler.asp?act=gizle&link_id="&rs("link_id")&"><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=linkler.asp?act=goster&link_id="&rs("link_id")&"><img src=""../images/yayinda_degil.png"" border=0></a>"
end if%></td>
                  <td align="center"><%=rs("link_gosterim")%></td>
                                 
                  <td align="center"><% Response.Write "<a href=linkler.asp?act=duzenle&link_id="&rs("link_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=linkler.asp?act=sil&link_id="&rs("link_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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
sub ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/linkler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Link Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="linkler.asp?act=ekle_bitir">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" height="25" bgcolor="#333333" class="style2">Yeni Link Ekle</td>
                  </tr>
                <tr>
                  <td height="25" bgcolor="#FFFFFF" class="style2"><table width="100%" border="0" cellpadding="3" cellspacing="3">
                    <tr align="left">
                      <td><table width="100%" border="0" cellpadding="0" cellspacing="0">

                        <tr bgcolor="#FFFFFF">
                          <td width="42%" height="25"><div align="right" class="style3">Baþlýk</div></td>
                          <td width="1%"><span class="style3">:</span></td>
                          <td width="57%"><input name="link_adi" type="text" class="inputbox2" id="link_adi" size="60" /></td>
                        </tr>
                        <tr bgcolor="fbe8a6">
                          <td height="25" bgcolor="fbe8a6"><div align="right" class="style3">Açýklama</div></td>
                          <td><span class="style3">:</span></td>
                          <td><textarea name="link_aciklama" cols="60" rows="3" class="inputbox2" id="link_aciklama"></textarea></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td height="25"><div align="right" class="style3">Link</div></td>
                          <td><span class="style3">:</span></td>
                          <td bgcolor="#FFFFFF" class="style10">http://
                              <input name="link_url" type="text" class="inputbox2" id="link_url" size="50" />                              </td>
                        </tr>


                      </table></td>
                    </tr>
                    <tr align="left">
                      <td width="48%" bgcolor="fbe8a6"><div align="center">
                        <input name="Submit2" type="submit" class="button" value="Ekle" onclick="tinyMCE.triggerSave();" />
                      </div></td>
                    </tr>
                  </table></td>
                </tr>
              </table>
              </form></td>
          </tr>
        </table>
<%
end sub

sub gizle
baglanti.Execute("UPDATE gop_linkler set link_onay='"& 0 &"' where link_id='" & request.querystring("link_id") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub goster
baglanti.Execute("UPDATE gop_linkler set link_onay='"& 1 &"' where link_id='" & request.querystring("link_id") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub sil
SQL="Delete From gop_linkler where link_id=" & request.querystring("link_id")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub ekle_bitir
link_adi = Request.Form("link_adi")
link_url = Request.Form("link_url")
link_aciklama = guvenlik(Request.Form("link_aciklama"))


SQL="insert into gop_linkler (link_adi,link_url,link_aciklama) values ('"&link_adi&"','"&link_url&"','"&link_aciklama&"')"
Baglanti.Execute(SQL)
Response.Redirect "linkler.asp"
end sub

sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/linkler.png" width="128" height="128" align="middle" /><span class="style6">Link Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_linkler where link_id=" & request.querystring("link_id")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="linkler.asp?act=duzenle_bitir&link_id=<%=rs("link_id")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Link Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25"><div align="right"><strong> ID</strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("link_id")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Baþlýk </strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><input name="link_adi" type="text" class="inputbox2" id="link_adi" value="<%=rs("link_adi")%>" size="60" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Açýklama</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><textarea name="link_aciklama" cols="60" rows="3" class="inputbox2" id="link_aciklama"><%=rs("link_aciklama")%></textarea></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>URL</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF">http://
                    <input name="link_url" type="text" class="inputbox2" id="link_url" value="<%=rs("link_url")%>" size="50" /></td>
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

sub duzenle_bitir
link_adi = Request.Form("link_adi")
link_url = Request.Form("link_url")
link_aciklama = guvenlik(Request.Form("link_aciklama"))

baglanti.Execute("UPDATE gop_linkler set link_adi='"&link_adi&"', link_url='"&link_url&"',link_aciklama='"&link_aciklama&"' where link_id='" & request.querystring("link_id") & "';")
response.Redirect "linkler.asp"
end sub

sub onay
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/linkler.png" width="128" height="128" align="middle" /><span class="style6"> Link Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="50" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Bilgi</strong></span><span class="style4"><strong></strong></span></td>
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
Set SQLToplam = baglanti.Execute("select count(link_id) from gop_linkler") 
TopKayit = SQLToplam(0)
set rs = baglanti.Execute("select * from gop_linkler where link_onay = '"& 0 &"' order by link_id asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)

if rs.eof or rs.bof then
response.Write "Link Bulunamadý"
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
<a href="http://<%=rs("link_url")%>" target="_blank"><strong><%=rs("link_adi")%></strong></a><br />
                    <%=rs("link_aciklama")%><br />
                    <br />
</td>
                  <td align="center"><%
if rs("link_onay") = "1" then
Response.Write "<a href=linkler.asp?act=gizle&link_id="&rs("link_id")&"><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=linkler.asp?act=goster&link_id="&rs("link_id")&"><img src=""../images/yayinda_degil.png"" border=0></a>"
end if%></td>
                  <td align="center"><%=rs("link_gosterim")%></td>
                                 
                  <td align="center"><% Response.Write "<a href=linkler.asp?act=duzenle&link_id="&rs("link_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=linkler.asp?act=sil&link_id="&rs("link_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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