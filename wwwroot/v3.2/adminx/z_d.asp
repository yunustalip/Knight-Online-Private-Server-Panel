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
		sira="Yazan"
	elseif siralama="2" then
		sira="yer"
	elseif siralam="3" then
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
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Ziyaretçi Defteri Mesajları</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
	<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo" align="center">
<form action="" method="get">
		<tr>
			<td align="right">
		<select name="siralama" size="1" class="alan">
        <option value="1"<%if sira="Yazan" then%> selected<%End if%>>İsme Göre</option>
        <option value="2"<%if sira="yer" then%> selected<%End if%>>Yere Göre</option>
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
if (Request.QueryString("zd"))="sil" then
id=request.querystring("id")
if isnumeric(id) then
data.Execute("DELETE FROM zd where id="&id&"")
end if
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if

if (Request.QueryString("mesaj"))="kayit" then
id=request.querystring("id")

mesaj=request.form("mesaj")
yer=request.form("yer")
islem=request.form("islem")
yazan=request.form("yazan")

if islem="Sil" then
if isnumeric(id) then
data.Execute("DELETE FROM zd where id="&id&"")
end if
else
if isnumeric(id) then
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from zd where id="&id&""
blgekle.Open SQL,data,1,3
if not blgekle.eof then
blgekle("mesaj")=mesaj
blgekle("yer")=yer
blgekle("yazan")=yazan
blgekle.update
end if
blgekle.Close
Set blgekle = Nothing
Response.Redirect "z_d.asp"
end if
End if
End if

set zd_msg = Server.CreateObject("Adodb.Recordset")
SQL = "Select * from zd where onay=0 order by "&sira&" "&t&""
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
		<td width="63" align="right"><font class="yazi">Ekleyen:</font></td>
		<td width="268"><input type="text" name="yazan" size="52" class="alan" value="<%=zd_msg("yazan")%>"></td>
		<td width="118" align="right"><font class="yazi">Yer:</font></td>
		<td width="794">
		<input type="text" name="yer" size="41" class="alan" value="<%=zd_msg("yer")%>"></td>
	</tr>
	<tr>
		<td width="63" align="right" valign="top"><font class="yazi">Mesajı:</font></td>
		<td width="1178" colspan="3">
		<textarea name="mesaj" rows="10" cols="120" class="alan"><%=zd_msg("mesaj")%></textarea></td>
	</tr>
	<tr>
		<td width="63">&nbsp;</td>
		<td width="1178" colspan="3"><font class="yazi">Tarih: <%=zd_msg("tarih")%>&nbsp;&nbsp;&nbsp;&nbsp;E-mail:<%=zd_msg("mail")%></font></td>
	</tr>
	<tr>
		<td width="63"></td>
		<td width="1178" colspan="3"><input type="submit" name="islem" value="Kaydet" class="dugme"><input type="submit" name="islem" value="Sil" class="dugme"></td>
	</tr>
</form>
</table>
<%zd_msg.movenext%>
<% next %> 
	<table width="99%" border="0" class="tablo" id="table1" cellpadding="0" style="border-collapse: collapse" align="center">
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