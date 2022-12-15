<!--#include file="admin_a.asp"-->
<style type="text/css">
<!--
.style6 {font-size: 24px}
.style7 {
	color: #CC0000;
	font-weight: bold;
}
-->
</style>
<%
islem = request.querystring("islem")
if islem = "yukle" then yukle
if islem = "uploading" then uploading
if islem = "kaldir" then kaldir
if islem = "" then default

sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000">
<span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Bileþen Yönetimi</span></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr align="left">
                <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                <td bgcolor="#333333"><span class="style4"><strong>Bileþen Adý</strong></span></td>
                <td bgcolor="#333333" align="center"><span class="style4"><strong>Mail</strong></span></td>
                <td bgcolor="#333333"><span class="style4"><strong>Yapýmcý</strong></span></td>
                <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
              </tr>
              <%
SQL ="SELECT * FROM gop_eklentiler order by id asc limit 0,999"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Bileþen Yok"
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
                <td ><%=rs("eklenti_adi")%></td>
                <td align="center"><a href="<%=rs("eklenti_web")%>" target="_blank"><%=rs("eklenti_mail")%></a></td>
                <td><%=rs("eklenti_yazar")%></td>
                <td><div align="center"><a href="?islem=kaldir&id=<%=rs("id")%>">Kaldýr</a></div></td>
              </tr>
              <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
            </table>
              <form action="bilesenler.asp?islem=uploading" method="post" enctype="multipart/form-data">
    <p class="style7">Yeni Bileþen Yükle      </p>
    <p>
      <input name="bilesen" type="file" class="inputbox2" id="bilesen" size="50">
      <input name="Submit" type="submit" class="button" value="Gönder">
      </p>
</form>            </td>
          </tr>
        </table>
<%
end sub
sub yukle
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Bileþen Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
kurulumadi = request.querystring("bilesen")
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
        if entry.tagName = "eklenti_adi" then
		eklenti_adi = entry.text
        elseif entry.tagname = "eklenti_k" then 
		eklenti_k = entry.text
		elseif entry.tagname = "eklenti_yazar" then 
		eklenti_yazar = entry.text
		elseif entry.tagname = "eklenti_mail" then 
		eklenti_mail = entry.text
		elseif entry.tagname = "eklenti_web" then 
		eklenti_web = entry.text
		elseif entry.tagname = "sqlcode" then 
		sqlcode = entry.text
		elseif entry.tagname = "sqlsil" then 
		sqlsil = entry.text
		elseif entry.tagname = "aspcode" then 
		aspcode = entry.text
		end if
    next
end If

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_eklentiler"
rs.open SQL,baglanti,1,3
rs.addnew
rs("eklenti_adi") = eklenti_adi
rs("eklenti_k") = eklenti_k
rs("eklenti_yazar") = eklenti_yazar
rs("eklenti_mail") = eklenti_mail
rs("eklenti_web") = eklenti_web
rs("eklenti") = aspcode
rs("eklenti_kaldir") = sqlsil
rs.update

if not sqlcode = "" then
Execute sqlcode
end if
Response.Write "<center><p><img src=""../images/ok.png""></p><p>Bileþen Baþarýyla Yüklendi</p></center>"
%>


            </td>
          </tr>
        </table>
<%
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

Set File = Upload.Files("bilesen")

If Not File Is Nothing Then
name=File.fileName
else
name=Null
end if
response.Write "<br><br><center><a href=""bilesenler.asp?islem=yukle&bilesen="&name&""">Yüklemeye Devam Et</a></center><br><br>"
%>


            </td>
          </tr>
        </table>
<%
end sub

sub kaldir

set sil=baglanti.execute("Select * from gop_eklentiler where id=" & request.querystring("id"))
if not sil("eklenti_kaldir") = "" then
Execute sil("eklenti_kaldir")
end if

SQL="Delete From gop_eklentiler where id=" & request.querystring("id")
Baglanti.Execute(SQL)
Response.Redirect "bilesenler.asp"
end sub

%>
<!--#include file="admin_b.asp"-->