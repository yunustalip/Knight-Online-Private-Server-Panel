<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "noonay" then onay
if islem = "onay" then noonay
if islem = "sil" then sil
if islem = "bekleyen" then bekleyen
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
                  <td bgcolor="#333333"><span class="style4"><strong>Yorum Yapýlan Mesaj</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Yorum Yapan</strong></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Yorumu Oku</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Onay</strong></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_yorumlar where yorum_onay = '"& 1 & "' order by yorum_id desc limit 0,100"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Yorum Yok"
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
                  <td ><% SQLvid ="SELECT * FROM gop_veriler where vid=" & rs("vid")
set vid = server.createobject("ADODB.Recordset")
vid.open SQLvid , Baglanti
if vid.eof or vid.bof then
response.Write "Veri Silindi"
else
response.Write vid("vbaslik")

vid.close
set vid = nothing
end if %></td>
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
                  <td align="center"><a href="#" onClick="MM_openBrWindow('yorum_oku.asp?yorum_id=<%=rs("yorum_id")%>','yorumlar','scrollbars=yes,width=350,height=400')">Oku</a>   </td>
                  <td align="center">
				  <%
if rs("yorum_onay") = "1" then
Response.Write "<a href=""yorumlar.asp?islem=onay&yorum_id="&rs("yorum_id")&"""><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=""yorumlar.asp?islem=noonay&yorum_id="&rs("yorum_id")&"""><img src=""../images/yayinda_degil.png"" border=0></a>"
end if
%>
	</td>
                  <td align="center"><% Response.Write "<a href=""yorumlar.asp?islem=sil&yorum_id="&rs("yorum_id")&"""><img src=""../images/sil.gif"" border=0></a>"%></td>
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
SQL="Delete From gop_yorumlar where yorum_id=" & request.querystring("yorum_id")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub onay
baglanti.Execute("UPDATE gop_yorumlar set yorum_onay='"& 1 &"' where yorum_id='" & request.querystring("yorum_id") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub noonay
baglanti.Execute("UPDATE gop_yorumlar set yorum_onay='"& 0 &"' where yorum_id='" & request.querystring("yorum_id") & "';")
Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub

sub bekleyen
%>

 <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/veriler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Veri Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Yorum Yapýlan Mesaj</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Yorum Yapan</strong></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Yorumu Oku</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Onay</strong></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_yorumlar where yorum_onay = '"& 0 & "' order by yorum_id desc limit 0,100"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Yorum Yok"
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
                  <td ><% SQLvid ="SELECT * FROM gop_veriler where vid=" & rs("vid")
set vid = server.createobject("ADODB.Recordset")
vid.open SQLvid , Baglanti
if vid.eof or vid.bof then
response.Write "Veri Silindi"
else
response.Write vid("vbaslik")

vid.close
set vid = nothing
end if %></td>
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
                  <td align="center"><a href="#" onClick="MM_openBrWindow('yorum_oku.asp?yorum_id=<%=rs("yorum_id")%>','yorumlar','scrollbars=yes,width=350,height=400')">Oku</a>   </td>
                  <td align="center">
				  <%
if rs("yorum_onay") = "1" then
Response.Write "<a href=""yorumlar.asp?islem=onay&yorum_id="&rs("yorum_id")&"""><img src=""../images/yayinda.png"" border=0></a>"
else
Response.Write "<a href=""yorumlar.asp?islem=noonay&yorum_id="&rs("yorum_id")&"""><img src=""../images/yayinda_degil.png"" border=0></a>"
end if
%>
	</td>
                  <td align="center"><% Response.Write "<a href=""yorumlar.asp?islem=sil&yorum_id="&rs("yorum_id")&"""><img src=""../images/sil.gif"" border=0></a>"%></td>
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
%>
<!--#include file="admin_b.asp"-->