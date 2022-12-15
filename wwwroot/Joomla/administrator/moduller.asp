<!--#include file="admin_a.asp"-->
<%
islem = request.querystring("islem")
if islem = "modul_noizin" then modul_noizin
if islem = "modul_izin" then modul_izin
if islem = "sira_yukari" then sira_yukari
if islem = "sira_asagi" then sira_asagi
if islem = "modul_pozisyon_sol" then modul_pozisyon_sol
if islem = "modul_pozisyon_sag" then modul_pozisyon_sag
if islem = "modul_pozisyon_orta_ust" then modul_pozisyon_orta_ust
if islem = "modul_pozisyon_orta_alt" then modul_pozisyon_orta_alt
if islem = "modul_pozisyon_orta_orta" then modul_pozisyon_orta_orta
if islem = "modul_sil" then modul_sil
if islem = "modul_ekle" then modul_ekle
if islem = "modul_manuel_ekle" then modul_manuel_ekle
if islem = "modul_duzenle" then modul_duzenle
if islem = "duzenle" then duzenle
if islem = "yukle" then yukle
if islem = "uploading" then uploading
if islem = "" then default

sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Modül Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Modul Adý</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Modül Pozisyon</strong></span></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Yapýmcý</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Sýra</strong></span></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Modül Yayýnla</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                </tr>
<%
SQL ="SELECT * FROM gop_modules order by modul_sira asc limit 0,999"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Modül Yok"
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
                  <td ><%=rs("modul_adi")%></td>
                  <td align="center">
<%
if rs("modul_yer") = "sag" then
Response.Write "<a href=moduller.asp?islem=modul_pozisyon_sol&modul_id="&rs("modul_id")&">" & rs("modul_yer") & "</a>"
elseif rs("modul_yer") = "sol" then
Response.Write "<a href=moduller.asp?islem=modul_pozisyon_orta_ust&modul_id="&rs("modul_id")&">" & rs("modul_yer") & "</a>"
elseif rs("modul_yer") = "orta_ust" then
Response.Write "<a href=moduller.asp?islem=modul_pozisyon_orta_alt&modul_id="&rs("modul_id")&">" & rs("modul_yer") & "</a>"
elseif rs("modul_yer") = "orta_alt" then
Response.Write "<a href=moduller.asp?islem=modul_pozisyon_orta_orta&modul_id="&rs("modul_id")&">" & rs("modul_yer") & "</a>"
elseif rs("modul_yer") = "orta_orta" then
Response.Write "<a href=moduller.asp?islem=modul_pozisyon_sag&modul_id="&rs("modul_id")&">" & rs("modul_yer") & "</a>"
end if

%>
                  
                  </td>
                  <td><a href="mailto:<%=rs("modul_mail")%>"><%=rs("modul_yazar")%></a></td>
                  <td align="center"><a href="moduller.asp?islem=sira_yukari&modul_id=<%=rs("modul_id")%>"><img src="../images/yukari.png" border="0"> </a><%=rs("modul_sira")%> <a href="moduller.asp?islem=sira_asagi&modul_id=<%=rs("modul_id")%>"><img src="../images/asagi.png" border="0"></a></td>
                                  <%
if rs("modul_izin") = "1" then
Response.Write "<td align=""center""><a href=moduller.asp?islem=modul_noizin&modul_id="&rs("modul_id")&"><img src=""../images/yayinda.png"" border=0></a></td>"
else
Response.Write "<td align=""center""><a href=moduller.asp?islem=modul_izin&modul_id="&rs("modul_id")&"><img src=""../images/yayinda_degil.png"" border=0></a></td>"
end if
%>
                  <td align="center"><% Response.Write "<a href=moduller.asp?islem=duzenle&modul_id="&rs("modul_id")&"><img src=""../images/duzenle.gif"" border=0></a>"%>
                    </td>
                  <td align="center"><%Response.Write "<a href=moduller.asp?islem=modul_sil&modul_id="&rs("modul_id")&"><img src=""../images/sil.gif"" border=0></a>"%></td>
                </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
<form action="moduller.asp?islem=uploading" method="post" enctype="multipart/form-data">
    <p class="style7">Yeni Modül Yükle      </p>
    <p>
      <input name="modules" type="file" class="inputbox2" id="modules" size="50">
      <input name="Submit" type="submit" class="button" value="Gönder">
      </p>
</form>
            </td>
          </tr>
        </table>
<%
end sub

sub modul_noizin
baglanti.Execute("UPDATE gop_modules set modul_izin='"& 0 &"' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub modul_izin
baglanti.Execute("UPDATE gop_modules set modul_izin='"& 1 &"' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub sira_yukari
Set rs = baglanti.Execute("Select * from gop_modules where modul_id=" & request.querystring("modul_id") & " ;")
if rs.eof or rs.bof then
Response.Redirect "hata2.asp"
else
baglanti.Execute("UPDATE gop_modules set modul_sira='"& rs("modul_sira") - 1 &"' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end if
end sub

sub sira_asagi
Set rs = baglanti.Execute("Select * from gop_modules where modul_id=" & request.querystring("modul_id") & " ;")
if rs.eof or rs.bof then
Response.Redirect "moduller.asp"
else
baglanti.Execute("UPDATE gop_modules set modul_sira='"& rs("modul_sira") + 1 &"' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end if
end sub

sub modul_pozisyon_sol
baglanti.Execute("UPDATE gop_modules set modul_yer='sol' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub modul_pozisyon_sag
baglanti.Execute("UPDATE gop_modules set modul_yer='sag' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub modul_pozisyon_orta_ust
baglanti.Execute("UPDATE gop_modules set modul_yer='orta_ust' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub modul_pozisyon_orta_alt
baglanti.Execute("UPDATE gop_modules set modul_yer='orta_alt' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub modul_pozisyon_orta_orta
baglanti.Execute("UPDATE gop_modules set modul_yer='orta_orta' where modul_id='" & request.querystring("modul_id") & "';")
response.Redirect "moduller.asp"
end sub

sub modul_sil
SQL="Delete From gop_modules where modul_id=" & request.querystring("modul_id")
Baglanti.Execute(SQL)
Response.Redirect "moduller.asp"
end sub


sub modul_ekle
modul_adi = Request.Form("modul_adi")
modul_icerik = Request.Form("modul_icerik")
modul_yer = Request.Form("modul_yer")
modul_yazar = Request.Form("modul_yazar")
modul_mail = Request.Form("modul_mail")


Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_modules"
rs.open SQL,baglanti,1,3
rs.addnew
rs("modul_adi") = modul_adi
rs("modul_icerik") = modul_icerik
rs("modul_yer") = modul_yer
rs("modul_yazar") = modul_yazar
rs("modul_mail") = modul_mail
rs.update

Response.Redirect "moduller.asp"
end sub

sub modul_duzenle
gelen_modul = Request.Form("modul_icerik")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_modules where modul_id='"& request.QueryString("modul_id")&"'"
rs.open SQL,baglanti,1,3

rs("modul_icerik") = gelen_modul

rs.update
Response.Redirect "moduller.asp"

end sub

sub uploading
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Modül Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
Set Upload = Server.CreateObject("Persits.Upload")  
Upload.OverwriteFiles = False ' True yapýlýrsa ayný isimdeki dosya üzerine yazar.
Upload.IgnoreNoPost = True
Upload.Save server.MapPath("../modules")&"\"

Set File = Upload.Files("modules")

If Not File Is Nothing Then
name=File.fileName
else
name=Null
end if
response.Write "<a href=moduller.asp?islem=yukle&modul="&name&">Yüklemeye Devam Et</a>"
%>


            </td>
          </tr>
        </table>
<%
end sub

sub yukle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Modül Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
kurulumadi = request.querystring("modul")
set xmlDoc = createObject("Microsoft.XMLDOM")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
dosya = Server.MapPath("../modules/"&kurulumadi&"")
xmlDoc.load (dosya)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatasý: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "moduladi" then
		moduladi = entry.text
        elseif entry.tagname = "aspcode" then 
		aspcode = entry.text
		elseif entry.tagname = "modulyazar" then 
		modulyazar = entry.text
		elseif entry.tagname = "mail" then 
		mail = entry.text
		elseif entry.tagname = "modulizin" then 
		modulizin = entry.text
		elseif entry.tagname = "modulsira" then 
		modulsira = entry.text
		elseif entry.tagname = "modulyer" then 
		modulyer = entry.text
		end if
    next
end If

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_modules"
rs.open SQL,baglanti,1,3
rs.addnew
rs("modul_adi") = moduladi
rs("modul_icerik") = aspcode
rs("modul_yazar") = modulyazar
rs("modul_mail") = mail
rs("modul_izin") = modulizin
rs("modul_sira") = modulsira
rs("modul_yer") = modulyer
rs.update
Response.Write "<center><p><img src=""../images/ok.png""></p><p>Modül Baþarýyla Yüklendi</p></center>"
if not sqlcode = "" then
Execute sqlcode
end if
%>


            </td>
          </tr>
        </table>
<%
end sub
sub duzenle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Modül Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_modules where modul_id=" & request.querystring("modul_id")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="moduller.asp?islem=modul_duzenle&modul_id=<%=rs("modul_id")%>" method="post">
              <table width="100%" border="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Modül Düzenle</td>
                  </tr>
                <tr>
                  <td>ID</td>
                  <td>:</td>
                  <td><%=rs("modul_id")%></td>
                </tr>
                <tr>
                  <td>Modül Adý </td>
                  <td>:</td>
                  <td><%=rs("modul_adi")%></td>
                </tr>
                <tr>
                  <td>Modül</td>
                  <td>:</td>
                  <td><textarea name="modul_icerik" cols="100" rows="20" class="inputbox2" id="modul_icerik"><%=rs("modul_icerik")%></textarea></td>
                </tr>
                <tr>
                  <td>Sýra</td>
                  <td>:</td>
                  <td><%=rs("modul_sira")%></td>
                </tr>
                <tr>
                  <td>Pozisyon</td>
                  <td>:</td>
                  <td><%=rs("modul_yer")%>
                  </td>
                </tr>
                <tr>
                  <td>Yazar</td>
                  <td>:</td>
                  <td><%=rs("modul_yazar")%></td>
                </tr>
                <tr>
                  <td>Mail</td>
                  <td>:</td>
                  <td><%=rs("modul_mail")%></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
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
sub modul_manuel_ekle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /> <span class="style6">Modül Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>            <td height="20"><form id="form1" name="form1" method="post" action="moduller.asp?islem=modul_ekle">
              <table width="100%" border="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Yeni Modül Ekle</td>
                  </tr>
                <tr>
                  <td>Modül Adý</td>
                  <td>:</td>
                  <td><input name="modul_adi" type="text" class="inputbox" id="modul_adi" /></td>
                </tr>
                <tr>
                  <td>Modül Kodu</td>
                  <td>:</td>
                  <td><textarea name="modul_icerik" cols="75" rows="20" class="inputbox2" id="modul_icerik"></textarea></td>
                </tr>
                <tr>
                  <td>Pozisyon</td>
                  <td>:</td>
                  <td><select name="modul_yer" class="inputbox" id="modul_yer">
                    <option value="sol">Sol</option>
                    <option value="sag">Sað</option>
                    <option value="orta_ust">Orta Üst</option>
                    <option value="orta_alt">Orta Alt</option>
                  </select>                  </td>
                </tr>
                <tr>
                  <td>Yazan</td>
                  <td>:</td>
                  <td><input name="modul_yazar" type="text" class="inputbox" id="modul_yazar" /></td>
                </tr>
                <tr>
                  <td>Mail</td>
                  <td>:</td>
                  <td><input name="modul_mail" type="text" class="inputbox" id="modul_mail" /></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><input name="Submit" type="submit" class="button" value="Ekle" /></td>
                </tr>
              </table>
              <p>&nbsp;</p>
              </form></td>
          </tr>
        </table>
<%
end sub
%>
<!--#include file="admin_b.asp"-->