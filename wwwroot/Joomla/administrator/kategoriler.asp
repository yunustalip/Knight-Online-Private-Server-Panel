<!--#include file="admin_a.asp"-->
<% 
act = request.querystring("act")
if act = "gizle" then
call gizle
elseif act = "goster" then
call goster
elseif act = "sira_yukari" then
call sira_yukari
elseif act = "sira_asagi" then
call sira_asagi
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
elseif act = "" then
call default
end if

sub default
%>

<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/kategori.png" width="128" height="128" align="middle" /><span class="style6"> Kategori Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20" bgcolor="e58e4d"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Kategori Adý</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>ID</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Sýra</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Linki Göster</strong></span></td>
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

SQL ="SELECT * FROM gop_anakat order by ankorder asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Kategori Yok"
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
                  <td ><%=rs("ankadi")%></td>
                  <td align="center"><%=rs("ankid")%></td>
                  <td align="center"><a href="kategoriler.asp?act=sira_yukari&ankid=<%=rs("ankid")%>"><img src="../images/yukari.png" border="0"> </a><%=rs("ankorder")%> <a href="kategoriler.asp?act=sira_asagi&ankid=<%=rs("ankid")%>"><img src="../images/asagi.png" border="0"></a></td>
                                  <%
if rs("ankgoster") = "1" then
Response.Write "<td align=""center""><a href=""kategoriler.asp?act=gizle&ankid="&rs("ankid")&"""><img src=""../images/yayinda.png"" border""=0""></a></td>"
else
Response.Write "<td align=""center""><a href=""kategoriler.asp?act=goster&ankid="&rs("ankid")&"""><img src=""../images/yayinda_degil.png"" border=""0""></a></td>"
end if
%>
                  <td align="center"><% Response.Write "<a href=""kategoriler.asp?act=duzenle&ankid="&rs("ankid")&"""><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=kategoriler.asp?act=sil&ankid="&rs("ankid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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
              </div>            </td>
          </tr>
        </table>
<%
end sub
sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/kategori.png" width="128" height="128" align="middle" /><span class="style6"> Kategori Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_anakat where ankid=" & request.querystring("ankid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="kategoriler.asp?act=duzenle_bitir&ankid=<%=rs("ankid")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Kategori Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="40%" height="25"><div align="right"><strong>ID</strong></div></td>
                  <td width="2%"><div align="right"><strong>:</strong></div></td>
                  <td width="58%"><%=rs("ankid")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Kategori Adý </strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td><input name="ankadi" type="text" class="inputbox" id="ankadi" value="<%=rs("ankadi")%>" maxlength="50" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Kategori Taným</strong></div></td>
                  <td><div align="right"><strong>:</strong></div></td>
                  <td><textarea name="ankbilgi" cols="45" rows="3" class="inputbox2" id="ankbilgi"><%=rs("ankbilgi")%></textarea></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25">&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><input name="Submit" type="submit" class="button" value="Düzenle" /></td>
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
ankadi = Request.Form("ankadi")
ankbilgi = Request.Form("ankbilgi")

baglanti.Execute("UPDATE gop_anakat set ankadi='"&ankadi&"', ankbilgi='"&ankbilgi&"' where ankid='" & request.querystring("ankid") & "';")
response.Redirect "kategoriler.asp"
end sub

sub gizle
baglanti.Execute("UPDATE gop_anakat set ankgoster='"& 0 &"' where ankid='" & request.querystring("ankid") & "';")
response.Redirect "kategoriler.asp"
end sub

sub goster
baglanti.Execute("UPDATE gop_anakat set ankgoster='"& 1 &"' where ankid='" & request.querystring("ankid") & "';")
response.Redirect "kategoriler.asp"
end sub


sub sira_yukari
Set rs = baglanti.Execute("Select * from gop_anakat where ankid=" & request.querystring("ankid") & " ;")
if rs.eof or rs.bof then
Response.Redirect "hata2.asp"
else
baglanti.Execute("UPDATE gop_anakat set ankorder='"& rs("ankorder") - 1 &"' where ankid='" & request.querystring("ankid") & "';")
response.Redirect "kategoriler.asp"
end if
end sub

sub sira_asagi
Set rs = baglanti.Execute("Select * from gop_anakat where ankid=" & request.querystring("ankid") & " ;")
if rs.eof or rs.bof then
Response.Redirect "hata2.asp"
else
baglanti.Execute("UPDATE gop_anakat set ankorder='"& rs("ankorder") + 1 &"' where ankid='" & request.querystring("ankid") & "';")
response.Redirect "kategoriler.asp"
end if
end sub

sub sil
SQL="Delete From gop_anakat where ankid=" & request.querystring("ankid")
Baglanti.Execute(SQL)
Response.Redirect "kategoriler.asp"
end sub

sub ekle_bitir
ankadi = Request.Form("ankadi")
ankbilgi = Request.Form("ankbilgi")


SQL="insert into gop_anakat (ankadi,ankbilgi) values ('"&ankadi&"','"&ankbilgi&"')"
Baglanti.Execute(SQL)
Response.Redirect "kategoriler.asp"
end sub

sub ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/kategori.png" width="128" height="128" align="middle" /> <span class="style6">Kategori Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="kategoriler.asp?act=ekle_bitir">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Yeni Kategori Ekle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="41%" height="25"><div align="right"><strong>Kategori Adý </strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="58%" bgcolor="fbe8a6"><input name="ankadi" type="text" class="inputbox" id="ankadi" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25" bgcolor="#FFFFFF"><div align="right"><strong>Kategori Taným</strong></div></td>
                  <td bgcolor="#FFFFFF"><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><textarea name="ankbilgi" cols="45" rows="3" class="inputbox2" id="ankbilgi"></textarea></td>
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
%>
        <!--#include file="admin_b.asp"-->