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
if islem = "duzenle" then
call duzenle
elseif islem = "guncelle" then
call guncelle
elseif islem = "yukle" then
call yukle
elseif islem = "yukle2" then
call yukle2
elseif islem = "sil" then
call sil
elseif islem = "" then
call default
end if
sub default
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/dil.png" alt="" width="128" height="128" align="middle" /><span class="style6"> Dil Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr align="left">
                  <td bgcolor="#333333" class="style4"><div align="center"><strong>#</strong></div></td>
                  <td bgcolor="#333333"><span class="style4"><strong>Dil Adý</strong></span></td>
                  <td align="center" bgcolor="#333333" class="style4"><strong>Yapýmcý</strong></td>
                  <td bgcolor="#333333" align="center"><span class="style4"><strong>Mail</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Düzenle</strong></span></td>
                  <td align="center" bgcolor="#333333"><span class="style4"><strong>Sil</strong></span></td>
                  </tr>
<%
SQL ="SELECT * FROM gop_language order by lang_id asc limit 0,100"
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
if rs.eof or rs.bof then
response.Write "Dil bulunamadý"
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
                  <td ><%=rs("lang_adi")%></td>
                  <td align="center"><%=rs("lang_yazar")%></td>
                  <td align="center"><%=rs("lang_yazar")%></td>
                  <td align="center"><% Response.Write "<a href=""diller.asp?islem=duzenle&lang_id="&rs("lang_id")&"""><img src=""../images/duzenle.gif"" border=""0""></a>"%>
                  </td>
                  <td align="center"><% Response.Write "<a href=""diller.asp?islem=sil&lang_id="&rs("lang_id")&"""><img src=""../images/sil.gif"" border=""0""></a>"%>                    </td>
                  </tr>
                <%
rs.movenext
next
rs.close
set rs = nothing
end if
%>
              </table>
<form action="diller.asp?islem=yukle" method="post" enctype="multipart/form-data">
    <p class="style7">Yeni Dil Yükle      </p>
    <p>
      <input name="diller" type="file" class="inputbox2" id="diller" size="50">
      <input name="Submit" type="submit" class="button" value="Gönder">
      </p>
</form>
            </td>
          </tr>
        </table>
 <% 
 end sub
 sub duzenle
 %>
 <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;</strong></span><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Dil Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr><%
SQL ="SELECT * FROM gop_language where lang_id=" & request.querystring("lang_id")
set rs = server.createobject("ADODB.Recordset")
rs.open SQL , Baglanti
%>
            <td height="20"><form action="diller.asp?islem=guncelle&lang_id=<%=rs("lang_id")%>" method="post">
              <table width="100%" border="0">
                <tr>
                  <td height="25" colspan="3" bgcolor="#333333" class="style2">Dil Düzenle</td>
                  </tr>
                <tr>
                  <td>ID</td>
                  <td>:</td>
                  <td><%=rs("lang_id")%></td>
                </tr>
                <tr>
                  <td>Dil Adý </td>
                  <td>:</td>
                  <td><%=rs("lang_adi")%></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><textarea name="language" cols="80" rows="30" class="inputbox2" id="language"><%=rs("language")%></textarea></td>
                </tr>
                <tr>
                  <td>Yazar</td>
                  <td>:</td>
                  <td><%=rs("lang_yazar")%></td>
                </tr>
                <tr>
                  <td>Mail</td>
                  <td>:</td>
                  <td><%=rs("lang_mail")%></td>
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
		sub guncelle
		Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_language where lang_id='"& request.QueryString("lang_id")&"'"
rs.open SQL,baglanti,1,3

rs("language") = dilkontrol(Request.Form("language"))
rs.update

Response.Redirect request.ServerVariables("HTTP_REFERER")
end sub
sub yukle
		%>
        <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Dil Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
Set Upload = Server.CreateObject("Persits.Upload")  
Upload.OverwriteFiles = False ' True yapýlýrsa ayný isimdeki dosya üzerine yazar.
Upload.IgnoreNoPost = True
Upload.Save server.MapPath("../modules")&"\"

Set File = Upload.Files("diller")

If Not File Is Nothing Then
name=File.fileName
else
name=Null
end if
response.Write "<br><br><center><a href=diller.asp?islem=yukle2&dil="&name&">Yüklemeye Devam Et</a></center><br><br>"
%>


            </td>
          </tr>
        </table>
        <% 
		end sub
		sub yukle2
	%>
    <table width="100%" border="0" cellpadding="2" cellspacing="2">
          <tr>
            <td height="20" bgcolor="#CC0000"><span class="style4"><strong>&nbsp;<img src="../images/modul.png" width="128" height="128" align="middle" /><span class="style6"> Dil Yönetimi</span></strong></span><span class="style4 style6"><strong></strong></span></td>
          </tr>
          <tr>
            <td height="20"><%
kurulumadi = request.querystring("dil")
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
        if entry.tagName = "diladi" then
		diladi = entry.text
        elseif entry.tagname = "aspcode" then 
		aspcode = entry.text
		elseif entry.tagname = "dilyazar" then 
		dilyazar = entry.text
		elseif entry.tagname = "mail" then 
		mail = entry.text
		end if
    next
end If

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_language"
rs.open SQL,baglanti,1,3
rs.addnew
rs("lang_adi") = diladi
rs("language") = aspcode
rs("lang_yazar") = dilyazar
rs("lang_mail") = mail
rs.update
Response.Write "<center><p><img src=""../images/ok.png""></p><p>Dil Baþarýyla Yüklendi</p></center>"
%>


            </td>
          </tr>
        </table>
        <%
		end sub
		sub sil
SQL="Delete From gop_language where lang_id=" & request.querystring("lang_id")
Baglanti.Execute(SQL)
Response.Redirect request.ServerVariables("HTTP_REFERER")
		end sub
		
		%>
        
        <!--#include file="admin_b.asp"-->