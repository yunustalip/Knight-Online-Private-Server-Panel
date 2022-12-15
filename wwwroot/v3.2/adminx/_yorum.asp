<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<!--#include file="../filtre.asp"-->
<body background="images/arka.gif">
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
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">
<%
id=Request.QueryString("id")
if not id="" then
if not isnumeric(id)=false then
set ktg = Server.CreateObject("ADODB.Recordset")
SQL = "Select id,konu from blog where id = "&id&""
ktg.open SQL,data,1,3
if not ktg.eof then
%>
<font class="baslik"><b><%=ktg("konu")%></b></font><font class="yazi">, konusuna ait</font><%
end if
ktg.close : set ktg = nothing
end if
end if
%><%if id="0" then%>Hakkımda Yapılan<%end if%> Yorumlar Listeleniyor</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
	<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="" method="get">
		<tr>
			<td align="right">
		<select name="siralama" size="1" class="alan">
        <option value="1"<%if sira="Ekleyen" then%> selected<%End if%>>İsme Göre</option>
        <option value="3"<%if sira="id" then%> selected<%End if%>>Tarihe Göre</option>
        </select><select name="tip" size="1" class="alan">
        <option value="1"<%if t="desc" then%> selected<%End if%>>Artan</option>
        <option value="2"<%if t="asc" then%> selected<%End if%>>Azalan</option>
        </select><select name="sayi" size="1" class="alan">
        <option value="1"<%if s="25" then%> selected<%End if%>>25</option>
        <option value="2"<%if s="50" then%> selected<%End if%>>50</option>
        <option value="3"<%if s="100" then%> selected<%End if%>>100</option>
        </select><input type="submit" value="Sırala" class="dugme">
			</td>
		</tr>
</form>
	</table>
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

blgekle("Ekleyen")=ekleyen
blgekle("yorum")=yorum

blgekle.update
blgekle.Close
Set blgekle = Nothing
Response.Redirect "_yorum.asp"
End if
End if
set zd_msg = Server.CreateObject("Adodb.Recordset")
if isnumeric(id)=false or id="" then
SQL = "Select * from yorum order by "&sira&" "&t&""
else
SQL = "Select * from yorum where blog_id="&id&" order by "&sira&" "&t&""
end if
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
<table border="0" width="99%" id="table1" cellpadding="0" style="border-collapse: collapse" class="tablo" align="center"<%if zd_msg("onay")="1" then%> bgcolor="#FFFFDD"<% End if %>>
<form action="?mesaj=kayit&id=<%=zd_msg("id")%>" method="POST">
	<tr>
		<td width="63" align="right">
		<font class="yazi">Ekleyen:</font></td>
		<td width="320"><input type="text" name="ekleyen" size="52" class="alan" value="<%=zd_msg("ekleyen")%>"></td>
	</tr>
	<tr>
		<td width="63" align="right" valign="top"><font class="yazi">Yorumu:</font></td>
		<td width="1178" colspan="2">
		<textarea name="yorum" rows="10" cols="120" class="alan"><%=zd_msg("yorum")%></textarea></td>
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
	<table width="99%" border="0" class="tablo" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td colspan="3" align="center"><font class="yazi">Toplam <%=adet%> kayıt, <%=sayfa_sayisi%> Sayfada Gösterilmektedir.</font></td>
			</tr>
		<tr>
			<td align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?sayfa=1&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""ilk sayfa"">««</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""Önceki"">«</a></b> "
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
Response.Write "<b><a href=""?sayfa=" & j & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""">" & j & "</a></b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""Sonraki"">»</a></b> "
Response.Write "<b><a href=""?sayfa=" & sayfa_sayisi & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""Son Sayfa"">»»</a></b>"
End If
zd_msg.close : set zd_msg = nothing
%>
			</td>
		</tr>
	</table>
</div>
<% End if %>
<% Else %>
<font class="yazi"><center>Kayıt Bulunamadı</center></font>
<% End if %>
<% end if %>