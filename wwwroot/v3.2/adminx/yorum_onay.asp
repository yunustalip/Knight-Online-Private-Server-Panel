<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<body background="images/arka.gif">
<!--#include file="db.asp"-->
<!--#include file="../filtre.asp"-->
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Onay Bekleyen Yorumlar</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table border="0" width="99%" id="table1" style="border-collapse: collapse" align="center" class="tablo">
	<tr>
		<td class="ust" align="right">
		<a href="?tumunu=onayla" onclick="return confirm('T�m� Onaylanacak\nEminmisin?');">T�m�n� Onayla</a> - <a href="?tumunu=sil" onclick="return confirm('T�m� Silinecek\nEminmisin?');">T�m�n� Sil</a></td>
	</tr>
</table>
<div align="center">
	<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="" method="get">
		<tr>
			<td align="right">
		<select name="siralama" size="1" class="alan">
        <option value="1"<%if sira="Ekleyen" then%> selected<%End if%>>�sme G�re</option>
        <option value="3"<%if sira="id" then%> selected<%End if%>>Tarihe G�re</option>
        </select><select name="tip" size="1" class="alan">
        <option value="1"<%if t="desc" then%> selected<%End if%>>Artan</option>
        <option value="2"<%if t="asc" then%> selected<%End if%>>Azalan</option>
        </select><select name="sayi" size="1" class="alan">
        <option value="1"<%if s="25" then%> selected<%End if%>>25</option>
        <option value="2"<%if s="50" then%> selected<%End if%>>50</option>
        <option value="3"<%if s="100" then%> selected<%End if%>>100</option>
        </select><input type="submit" value="S�rala" class="dugme">
			</td>
		</tr>
</form>
	</table>
</div>
<%
siralama=filtre(Request.QueryString("siralama"))
tip=filtre(Request.QueryString("tip"))
sayi=filtre(Request.QueryString("sayi"))

	if siralama="1" then
		sira="Ekleyen"
	elseif siralama="3" then
		sira="id"
	else
		sira="id"
	end if

	if tip="1" then
		t="desc"
	elseif tip="2" then
		t="asc"
	else
		t="desc"
	end if

	if sayi="1" then
		s="25"
	elseif sayi="2" then
		s="50"
	elseif sayi="3" then
		s="100"
	else
		s="25"
	end if
%>
<%
if (Request.QueryString("mesaj"))="kayit" then
id=request.querystring("id")

yorum=request.form("yorum")
islem=request.form("islem")
ekleyen=request.form("ekleyen")

if islem="Sil" then
data.Execute("DELETE FROM yorum where id="&id&"")
else

set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from yorum where id="&id&""
blgekle.Open SQL,data,1,3
if not blgekle.eof then
blgekle("Ekleyen")=ekleyen
blgekle("yorum")=yorum
blgekle("onay")=0

blgekle.update
end if
blgekle.Close
Set blgekle = Nothing
response.redirect "yorum_onay.asp"
End if
End if

if (Request.QueryString("tumunu"))="sil" then
data.Execute("DELETE FROM yorum where onay=1")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if

set zd_msg = Server.CreateObject("Adodb.Recordset")
SQL = "Select * from yorum where onay=1 order by "&sira&" "&t&""
zd_msg.open SQL,data,1,3

adet = zd_msg.recordcount
if not zd_msg.eof then

sayfa = Request.QueryString("sayfa")
    if isnumeric(sayfa)=false then
        Response.redirect "index.asp"
    Else
if sayfa="" then sayfa=1
zd_msg.pagesize = (s)
sayfa_sayisi = zd_msg.pagecount
if Cint(sayfa)<1 then sayfa=1
if Cint(sayfa_sayisi) < Cint(sayfa) then sayfa=sayfa_sayisi
zd_msg.absolutepage = sayfa
for i=1 to zd_msg.pagesize
if zd_msg.eof then
exit for
end if
%>
<table border="0" width="99%" id="table1" cellpadding="0" style="border-collapse: collapse" class="tablo" align="center">
<form action="?mesaj=kayit&id=<%=zd_msg("id")%>" method="POST">
	<tr>
		<td width="63">
		<p align="right"><font class="yazi">Ekleyen:</font></td>
		<td width="320"><input type="text" name="ekleyen" size="52" class="alan" value="<%=zd_msg("ekleyen")%>"></td>
	</tr>
	<tr>
		<td width="63" align="right" valign="top"><font class="yazi">Mesaj�:</font></td>
		<td width="1178" colspan="2">
		<textarea name="yorum" rows="8" cols="97" class="alan"><%=zd_msg("yorum")%></textarea></td>
	</tr>
	<tr>
		<td width="63">&nbsp;</td>
<%
set ktg = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from blog where id = "&zd_msg("blog_id") &""
ktg.open SQL,data,1,3
if ktg.eof then
else
end if
if zd_msg("blog_id")="0" then
link="../hakkimda.asp"
else
link="../"&SEOLink(zd_msg("blog_id"))
end if
%>
		<td width="1178" colspan="2"><font class="yazi">Tarih: <%=zd_msg("tarih")%>&nbsp;&nbsp;&nbsp;&nbsp;Adres:<a href="<%=link%>"><%=link%></a></font></td>
<% ktg.close : set ktg = nothing %>
	</tr>
	<tr>
		<td width="63"></td>
		<td width="1178" colspan="2"><input type="submit" name="islem" value="Kaydet" class="dugme"><input type="submit" name="islem" value="Sil" class="dugme"></td>
	</tr>
</form>
</table>
<%zd_msg.movenext%>
<% next %> 
	</table>
	<table width="99%" border="0" class="tablo" id="table1" cellpadding="0" style="border-collapse: collapse" align="center">
		<tr>
			<td colspan="3" align="center"><font class="yazi">Toplam <%=adet%> kay�t, <%=sayfa_sayisi%> Sayfada G�sterilmektedir.</font></td>
			</tr>
		<tr>
			<td align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?sayfa=1&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""ilk sayfa"">��</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""�nceki"">�</a></b> "
End If
If sayfa + 10 > sayfa_sayisi Then
b = sayfa_sayisi 
Elseif sayfa_sayisi < 10 Then
sayfa_sayisi = 1
Else
b = sayfa + 10
End If
If sayfa < 10 Then
c = 1
Else
c = sayfa - 10
End If
if c < 1 then 
c = 1
end if
For j = c To b
If j = CInt(sayfa) Then
Response.Write "<font class=""yazi"">[<b>" & j & "</b>] </font>"
Else
Response.Write "<b><a href=""?sayfa=" & j & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""">" & j & "</a></b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""Sonraki"">�</a></b> "
Response.Write "<b><a href=""?sayfa=" & sayfa_sayisi & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""Son Sayfa"">��</a></b>"
End If
zd_msg.close : set zd_msg = nothing
%>
			</td>
		</tr>
	</table>
</div>
<% End if %>
<% Else %>
<font class="yazi"><center>Kay�t Bulunamad�</center></font>
<% End if %>
<% end if %>