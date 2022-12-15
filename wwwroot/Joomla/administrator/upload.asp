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
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<!--#include file="kontrol.asp"-->
<%
islem = request.querystring("islem")
if islem = "gonder" then
call gonder
elseif islem = "ekle" then
call ekle
elseif islem = "veri_resimekle" then
call veri_resimekle
elseif islem = "veri_resimgonder" then
call veri_resimgonder
elseif islem = "resim_boyutlandir" then
call resim_boyutlandir
elseif islem = "vresim_boyutlandir" then
call vresim_boyutlandir
elseif islem = "" then
call ana
end if
%>
<% sub ana
Response.Write "Geçersiz Ýstek"
 %>
<% end sub %>
<% sub gonder %>
<form action="upload.asp?islem=ekle" method="post" enctype="multipart/form-data">
    <input name="resresim" type="file" id="resresim">
    <input type="submit" name="Submit" value="Gönder">
  </form>
<% end sub

sub ekle
Set Upload = Server.CreateObject("Persits.Upload")  
Upload.OverwriteFiles = False ' True yapýlýrsa ayný isimdeki dosya üzerine yazar.
Upload.IgnoreNoPost = True
Upload.Save server.MapPath("../galeri")&"\"

Set File = Upload.Files("resresim")

If Not File Is Nothing Then
name=File.fileName
else
name=Null
end if
Response.Redirect "upload.asp?islem=resim_boyutlandir&resim_adi="&name

end sub

sub resim_boyutlandir

if ayar("aspjpeg") = "evet" then
%>

<script>
function kaydet()
{
window.opener.document.forms('form1').resresim.value=document.getElementById('F1').value
self.close()
}
</script>
<%
Set Jpeg = Server.CreateObject("Persits.Jpeg")
Jpeg.Open Server.MapPath("../galeri/"&Request.QueryString ("resim_adi"))
L = 100
H = 100
Jpeg.Width = L
Jpeg.Height = H
Jpeg.Save Server.MapPath("../galeri/thumbnail/"&Request.QueryString ("resim_adi"))
%>
<form method="POST" action="--WEBBOT-SELF--" enctype="multipart/form-data">
<p><input type="text" name="F1" value="../galeri/<%Response.Write(Request.QueryString ("resim_adi"))%>"></p>
<p><a href="javascript:kaydet()">onayla</a></p>
</form>
<%
else
%>
<form method="POST" action="--WEBBOT-SELF--" enctype="multipart/form-data">
<p><input type="text" name="F1" value="../galeri/<%Response.Write(Request.QueryString ("resim_adi"))%>"></p>
<p><a href="javascript:kaydet()">onayla</a></p>
</form>
<%
end if
end sub
%>

<% sub veri_resimgonder %>
<form action="upload.asp?islem=veri_resimekle" method="post" enctype="multipart/form-data">
    <input name="vresim" type="file" id="vresim">
    <input type="submit" name="Submit" value="Gönder">
  </form>
<% end sub

sub veri_resimekle
Set Upload = Server.CreateObject("Persits.Upload")  
Upload.OverwriteFiles = False ' True yapýlýrsa ayný isimdeki dosya üzerine yazar.
Upload.IgnoreNoPost = True
Upload.Save server.MapPath("../upload")&"\"

Set File = Upload.Files("vresim")

If Not File Is Nothing Then
name=File.fileName
else
name=Null
end if
Response.Redirect "upload.asp?islem=vresim_boyutlandir&resim_adi="&name

end sub

sub vresim_boyutlandir

if ayar("aspjpeg") = "evet" then
%>

<script>
function kaydet()
{
window.opener.document.forms('form1').vresim.value=document.getElementById('F1').value
self.close()
}
</script>
<%
Set Jpeg = Server.CreateObject("Persits.Jpeg")
Jpeg.Open Server.MapPath("../upload/"&Request.QueryString ("resim_adi"))
L = 100
H = 100
Jpeg.Width = L
Jpeg.Height = H
Jpeg.Save Server.MapPath("../upload/thumbnail/"&Request.QueryString ("resim_adi"))
%>
<form method="POST" action="--WEBBOT-SELF--" enctype="multipart/form-data">
<p><input type="text" name="F1" value="../upload/<%Response.Write(Request.QueryString ("resim_adi"))%>"></p>
<p><a href="javascript:kaydet()">Onayla</a></p>
</form>
<%
else
%>
<form method="POST" action="--WEBBOT-SELF--" enctype="multipart/form-data">
<p><input type="text" name="F1" value="../upload/<%Response.Write(Request.QueryString ("resim_adi"))%>"></p>
<p><a href="javascript:kaydet()">Onayla</a></p>
</form>
<%
end if
end sub
%>
<%
else
Response.Redirect "hata.asp"
end if
else
Response.Redirect "hata.asp"
end if
%>