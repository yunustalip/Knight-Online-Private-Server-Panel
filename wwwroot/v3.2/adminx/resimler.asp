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
		sira="isim"
	elseif siralama="2" then
		sira="hit"
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
		<td width="1215" background="images/bg.gif"><center>
<%
id=Request.QueryString("id")
if not id="" then
if not isnumeric(id)=false then
set ktg = Server.CreateObject("ADODB.Recordset")
SQL = "Select id,isim from galeri_kat where id = "&id&""
ktg.open SQL,data,1,3
if not ktg.eof then
%>
<font class="baslik"><b>&nbsp;&nbsp;&nbsp;&nbsp;<%=ktg("isim")%></b></font>, <font class="yazi">Kategorisine Ait</font>
<%
end if
ktg.close : set ktg = nothing
end if
end if
%>
		<font class="baslik">Resimler</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
	<div align="center">
	<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="" method="get">
		<tr>
			<td>
		<input type="hidden" value="<%=id%>" name="id">
		<select name="siralama" size="1" class="alan">
        <option value="1"<%if sira="isim" then%> selected<%End if%>>�sme G�re</option>
        <option value="2"<%if sira="id" then%> selected<%End if%>>Tarihe G�re</option>
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
<div align="center">
<table border="0" width="99%" id="table2" style="border-collapse: collapse" class="tablo">
	<tr>
		<td height="24" class="ust" width="24"><font class="yazi">ID</font></td>
		<td height="24" class="ust" width="1080"><font class="yazi">�sim</font></td>
		<td height="24" class="ust" width="116"><font class="yazi">��lem</font></td>
	</tr>
	
<% response.buffer = "true" %>

<%
if (Request.QueryString("resim"))="sil" then
id=request.querystring("id")
data.Execute("DELETE FROM galeri where id like '"&id&"'")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if

set zd_msg = Server.CreateObject("Adodb.Recordset")
if isnumeric(id)=false or id="" then
SQL = "Select * from galeri order by "&sira&" "&t&""
else
SQL = "Select * from galeri where kat_id="&id&" order by "&sira&" "&t&""
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
mode = 2
for i=1 to zd_msg.pagesize
if zd_msg.eof then
exit for
end if
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
%>
	<tr>
		<td width="24" height="20" class="<%=stil%>"><font class="yazi"><%=zd_msg("id")%></font></td>
		<td width="1080" height="20" class="<%=stil%>"><a href="resim_duz.asp?resim=duzenle&id=<%=zd_msg("id")%>"><%=zd_msg("isim")%></a></td>
		<td width="116" height="20" class="<%=stil%>"><a href="?resim=sil&id=<%=zd_msg("id")%>" onclick="return confirm('Silmek �stedi�inizden Eminmisiniz?');">Sil</a></td>
	</tr>
<%
zd_msg.movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
%>
<% next %> 
	</table>
</div>
<table width="99%" border="0" class="tablo" id="table1" cellpadding="0" style="border-collapse: collapse" align="center">
		<tr>
			<td colspan="3" align="center"><font class="yazi">Toplam <%=adet%> kay�t, <%=sayfa_sayisi%> Sayfada G�sterilmektedir.</font></td>
			</tr>
		<tr>
			<td align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?sayfa=1&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""ilk sayfa"">��</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""�nceki"">�</a></b> "
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
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""Sonraki"">�</a></b> "
Response.Write "<b><a href=""?sayfa=" & sayfa_sayisi & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&"&id="&id&""" title=""Son Sayfa"">��</a></b>"
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
<% End if %>