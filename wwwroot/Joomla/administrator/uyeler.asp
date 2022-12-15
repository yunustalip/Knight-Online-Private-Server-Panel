<!--#include file="admin_a.asp"-->
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
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/uyeler.png" width="128" height="128" align="middle" /><span class="style6"> Üye Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr bgcolor="e58e4d">
            <td height="20" bgcolor="e58e4d"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Üye Adý</strong></span></td>
                  <td bgcolor="#333333" align="center"><strong class="style4">ID</strong></td>
                  <td width="250" align="center" bgcolor="#333333"><span class="style4"><strong>E-Mail</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Son Giriþ</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Üyelik Tarihi</strong></span></td>
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

set rs = baglanti.Execute("select * from gop_uyeler order by uye_id asc LIMIT "& (listele*Sayfa)-(listele) & "," & listele)
Set SQLToplam = baglanti.Execute("select count(uye_id) from gop_uyeler") 
TopKayit = SQLToplam(0)
if rs.eof or rs.bof then
response.Write "Üye Yok"
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
                  <td ><%=rs("uye_adi")%></td>
                  <td align="center"><%=rs("uye_id")%></td>
                  <td ><%=rs("uye_mail")%></td>
                  <td align="center"><%=rs("uye_son_tarih")%></td>
                  <td align="center"><%=rs("uye_tarih")%></td>
                                 
                  <td align="center"><% Response.Write "<a href=uyeler.asp?islem=duzenle&uye_id="&rs("uye_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>                    </td>
                  <td align="center"><%Response.Write "<a href=uyeler.asp?islem=sil&uye_id="&rs("uye_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
              <!--#include file="sayfa.asp"-->              </td>
          </tr>
        </table>
        <%
		end sub
		sub duzenle
		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/uyeler.png" width="128" height="128" align="middle" /><span class="style6">Üye Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_uyeler where uye_id=" & request.querystring("uye_id")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="uyeler.asp?islem=duzenle_bitir&uye_id=<%=rs("uye_id")%>" method="post">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Üye Düzenle</td>
                  </tr>
                <tr bgcolor="fbe8a6">
                  <td width="42%" height="25"><div align="right"><strong> ID</strong></div></td>
                  <td width="1%"><strong>:</strong></td>
                  <td width="57%" bgcolor="fbe8a6"><%=rs("uye_id")%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Üye Adý </strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><input name="uye_adi" type="text" class="inputbox" id="uye_adi" value="<%=rs("uye_adi")%>" maxlength="50" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>E-Mail</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_mail" type="text" class="inputbox" id="uye_mail" value="<%=rs("uye_mail")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="25"><div align="right"><strong>Group</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="#FFFFFF"><select name="gid" class="inputbox" id="gid">
        <option value="3">Üye</option>
        <option value="1">Yönetici</option>
            </select></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Ýsim</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_isim" type="text" class="inputbox" id="uye_isim" value="<%=rs("uye_isim")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Soyisim</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_soyisim" type="text" class="inputbox" id="uye_soyisim" value="<%=rs("uye_soyisim")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Web Site</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_website" type="text" class="inputbox" id="uye_website" value="<%=rs("uye_website")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Ülke</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_ulke" type="text" class="inputbox" id="uye_ulke" value="<%=rs("uye_ulke")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Þehir</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_sehir" type="text" class="inputbox" id="uye_sehir" value="<%=rs("uye_sehir")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Msn</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_msn" type="text" class="inputbox" id="uye_msn" value="<%=rs("uye_msn")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Icq</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_icq" type="text" class="inputbox" id="uye_icq" value="<%=rs("uye_icq")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Aol</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_aol" type="text" class="inputbox" id="uye_aol" value="<%=rs("uye_aol")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Yahoo</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_yahoo" type="text" class="inputbox" id="uye_yahoo" value="<%=rs("uye_yahoo")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><span style="font-weight: bold">Skype</span></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_skype" type="text" class="inputbox" id="uye_skype" value="<%=rs("uye_skype")%>" size="60" maxlength="255" /></td>
                </tr>
                <tr bgcolor="fbe8a6">
                  <td height="25"><div align="right"><strong>Þifre</strong></div></td>
                  <td><strong>:</strong></td>
                  <td bgcolor="fbe8a6"><input name="uye_sifre" type="text" class="inputbox" id="uye_sifre" />
                    <span class="style7">Þifreyi deðiþtirmek istemiyorsanýz boþ býrakýnýz.</span></td>
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

uye_adi = Request.Form("uye_adi")
uye_sifre = md5(Request.Form ("uye_sifre"))
uye_mail = Request.Form("uye_mail")
gid = Request.Form("gid")
uye_isim = Request.Form("uye_isim")
uye_soyisim = Request.Form("uye_soyisim")
uye_website = Request.Form("uye_website")
uye_ulke = Request.Form("uye_ulke")
uye_sehir = Request.Form("uye_sehir")
uye_msn = Request.Form("uye_msn")
uye_icq = Request.Form("uye_icq")
uye_aol = Request.Form("uye_aol")
uye_yahoo = Request.Form("uye_yahoo")
uye_skype = Request.Form("uye_skype")

If request.form("uye_sifre")="" then
baglanti.Execute("UPDATE gop_uyeler set uye_adi='"&uye_adi&"', uye_mail='"&uye_mail&"', gid='"&gid&"', uye_isim='"&uye_isim&"', uye_soyisim='"&uye_soyisim&"', uye_website='"&uye_website&"', uye_ulke='"&uye_ulke&"', uye_sehir='"&uye_sehir&"', uye_msn='"&uye_msn&"', uye_icq='"&uye_icq&"', uye_aol='"&uye_aol&"', uye_yahoo='"&uye_yahoo&"', uye_skype='"&uye_skype&"' where uye_id='" & request.querystring("uye_id") & "';")
response.Redirect "uyeler.asp"
else
baglanti.Execute("UPDATE gop_uyeler set uye_adi='"&uye_adi&"', uye_mail='"&uye_mail&"', gid='"&gid&"', uye_isim='"&uye_isim&"', uye_soyisim='"&uye_soyisim&"', uye_website='"&uye_website&"', uye_ulke='"&uye_ulke&"', uye_sehir='"&uye_sehir&"', uye_msn='"&uye_msn&"', uye_icq='"&uye_icq&"', uye_aol='"&uye_aol&"', uye_yahoo='"&uye_yahoo&"', uye_skype='"&uye_skype&"', uye_sifre='"&uye_sifre&"' where uye_id='" & request.querystring("uye_id") & "';")
response.Redirect "uyeler.asp"
end if

end sub

sub ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/uyeler.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Üye Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="uyeler.asp?islem=ekle_bitir">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="100%" height="25" bgcolor="#333333" class="style2">Yeni Üye Ekle</td>
                  </tr>
                <tr>
                  <td height="25" bgcolor="#FFFFFF" class="style2"><table width="100%" border="0" cellpadding="3" cellspacing="3">
                    <tr align="left">
                      <td><table width="100%" border="0" cellpadding="0" cellspacing="0">

                        <tr bgcolor="#FFFFFF">
                          <td width="42%" height="25"><div align="right" class="style3">Üye Adý </div></td>
                          <td width="1%"><span class="style3">:</span></td>
                          <td width="57%"><input name="uye_adi" type="text" class="inputbox " id="uye_adi" maxlength="50" /></td>
                        </tr>
                        <tr bgcolor="fbe8a6">
                          <td height="25"><div align="right" class="style3">E-Mail</div></td>
                          <td><span class="style3">:</span></td>
                          <td><input name="uye_mail" type="text" class="inputbox " id="uye_mail" size="60" maxlength="255" /></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td height="25"><div align="right" class="style3">Group</div></td>
                          <td><span class="style3">:</span></td>
                          <td><select name="gid" class="inputbox style7" id="gid">
                              <option value="3">Üye</option>
                              <option value="1">Yönetici</option>
                          </select></td>
                        </tr>
                        <tr bgcolor="fbe8a6">
                          <td height="25"><div align="right" class="style7"><strong>Þifre</strong></div></td>
                          <td><span class="style7"><strong>:</strong></span></td>
                          <td><input name="uye_sifre" type="text" class="inputbox " id="uye_sifre" /> 
                            <span class="style9">Boþ býrakýrsanýz þifre 123456 olarak kayda geçecektir.</span></td>
                        </tr>

                      </table></td>
                    </tr>
                    <tr align="left">
                      <td width="48%"><div align="center">
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

sub ekle_bitir

uye_adi = Request.Form("uye_adi")
uye_sifre = md5(Request.Form ("uye_sifre"))
uye_mail = Request.Form("uye_mail")
gid = Request.Form("gid")
uye_tarih = tarih2
uye_son_tarih = tarih2

if Request.Form("uye_sifre")="" then
uye_sifre = md5("123456")
SQL="insert into gop_uyeler (uye_adi,uye_sifre,uye_mail,gid,uye_tarih,uye_son_tarih) values ('"&uye_adi&"','"&uye_sifre&"','"&uye_mail&"','"&gid&"','"&uye_tarih&"','"&uye_son_tarih&"')"
Baglanti.Execute(SQL)
Response.Redirect "uyeler.asp"
else
SQL="insert into gop_uyeler (uye_adi,uye_sifre,uye_mail,gid,uye_tarih,uye_son_tarih) values ('"&uye_adi&"','"&uye_sifre&"','"&uye_mail&"','"&gid&"','"&uye_tarih&"','"&uye_son_tarih&"')"
Baglanti.Execute(SQL)
Response.Redirect "uyeler.asp"
end if
end sub

sub sil
SQL="Delete From gop_uyeler where uye_id=" & request.querystring("uye_id")
Baglanti.Execute(SQL)

Response.Redirect "uyeler.asp"
end sub
%>
<!--#include file="admin_b.asp"-->