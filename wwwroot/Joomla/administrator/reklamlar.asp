<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "sil" then sil
if islem = "duzenle" then duzenle
if islem = "ekle" then ekle
if islem = "duzenle_bitir" then duzenle_bitir
if islem = "gizle" then gizle
if islem = "goster" then goster
if islem = "sil" then sil
if islem = "" then default

sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/reklamlar.png" width="128" height="128" align="middle" /><span class="style6"> Reklam Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20" bgcolor="e58e4d"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td width="40" bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Site Adý</strong></span></td>
                  <td width="468" align="center" bgcolor="#333333"><span class="style4"><strong>Banner</strong></span></td>
                  <td width="50" align="center" bgcolor="#333333"><span class="style4"><strong>Hit</strong></span></td>
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Gösterim</strong></span></td>
                  
                  <td width="80" align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td width="50" align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
listele = 25
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 

SQL ="SELECT * FROM gop_reklam order by hit asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele

Set SQLToplam = baglanti.Execute("select count(rid) from gop_reklam") 
TopKayit = SQLToplam(0)

set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Reklam Yok"
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
                  <td><a href="<%=rs("radi")%>" target="_blank"><%=rs("radi")%></a></td>
                  <td><img src="<%=rs("rresim")%>" width="468" height="60" border=0></td>
                  <td align="center"><%=rs("hit")%></td>
                  <td align="center">
<%
if rs("rgoster") = "1" then
Response.Write "<a href=reklamlar.asp?islem=gizle&rid="&rs("rid")&"><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=reklamlar.asp?islem=goster&rid="&rs("rid")&"><img src=""../images/yayinda_degil.png"" border=0></a>"
end if
%></td>
                  <td align="center"><% Response.Write "<a href=reklamlar.asp?islem=duzenle&rid="&rs("rid")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=reklamlar.asp?islem=sil&rid="&rs("rid")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
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
            <td height="20" bgcolor="#CC0000" class="style2">Reklam Ekle</td>
          </tr>
          <tr>
            <td height="30" bgcolor="#FFFFFF"><form id="form1" name="form1" method="post" action="reklamlar.asp?islem=ekle">
              Reklam Adý :
              <input name="radi" type="text" class="inputbox2" id="radi" size="30" />
Resim URL:
<input name="rresim" type="text" class="inputbox2" id="rresim" size="30" />
Link:
<input name="rlink" type="text" class="inputbox2" id="rlink" size="30" />
<input name="button" type="submit" class="button" id="button" value="Reklam Ekle" />

                        </form></td>
          </tr>
        </table>
<%
end sub

sub sil
SQL="Delete From gop_reklam where rid=" & request.querystring("rid")
Baglanti.Execute(SQL)
Response.Redirect "reklamlar.asp"
end sub

sub duzenle_bitir
radi = guvenlik(Request.Form("radi"))
rresim = Request.Form("rresim")
rlink = Request.Form("rlink")

baglanti.Execute("UPDATE gop_reklam set radi='"&radi&"', rresim='"&rresim&"', rlink='"&rlink&"' where rid='" & request.querystring("rid") & "';")
response.Redirect "reklamlar.asp"
end sub

sub ekle
radi = guvenlik(Request.Form("radi"))
rresim = Request.Form("rresim")
rlink = Request.Form("rlink")


SQL="insert into gop_reklam (radi,rresim,rlink) values ('"&radi&"','"&rresim&"','"&rlink&"')"
Baglanti.Execute(SQL)
Response.Redirect "reklamlar.asp"
end sub

sub goster
baglanti.Execute("UPDATE gop_reklam set rgoster='"& 1 &"' where rid='" & request.querystring("rid") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub gizle
baglanti.Execute("UPDATE gop_reklam set rgoster='"& 0 &"' where rid='" & request.querystring("rid") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/reklamlar.png" width="128" height="128" align="middle" /><span class="style6">Reklam Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_reklam where rid=" & request.querystring("rid")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="reklamlar.asp?islem=duzenle_bitir&rid=<%=rs("rid")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Reklam Düzenle</td>
                  </tr>
                <tr bgcolor="#FFFFFF">
                  <td width="42%" height="25"><div align="right"><strong>Reklam Adý</strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="57%" bgcolor="#FFFFFF"><input name="radi" type="text" class="inputbox2" id="radi" value="<%=rs("radi")%>" size="45" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Link</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="rlink" type="text" class="inputbox2" id="rlink" value="<%=rs("rlink")%>" size="45" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Resim URL</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="rresim" type="text" class="inputbox2" id="rresim" value="<%=rs("rresim")%>" size="45" /></td>
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