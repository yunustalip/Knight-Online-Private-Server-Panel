<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Etiket Düzenle</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<!--#include file="../filtre.asp"-->
<body background="images/arka.gif">
<%
if (Request.QueryString("etk"))="kaydet" then
id=Request.QueryString("id")
if not isnumeric(id)=false or id="" then
	etiket=Request.Form("etiket")
	if etiket="" then
	response.write "<center><font class=yazi>Etiket Alanı Boş Bırakılamaz</font></center>"
	response.end
	end if
set etkkaydet = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from etiket where id="&id&""
etkkaydet.Open SQL,data,1,3
if not etkkaydet.eof then
	etkkaydet("etiket")=Filtre(Trim(etiket))
	etkkaydet.update
end if
etkkaydet.close : set etkkaydet=nothing
Response.Redirect "etiket.asp"
end if
end if

if (Request.QueryString("Etiketi"))="Duzenle" then
id=Request.QueryString("id")
if not id="" then
if not isnumeric(id)=false then
set etiket = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from etiket where id = "&id&""
etiket.open SQL,data,1,3
if not etiket.eof then
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik"><%=etiket("etiket")%> - Etiket Düzenleme</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
<table border="0" width="99%" id="table3" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="?etk=kaydet&id=<%=id%>" method="post">
	<tr>
		<td width="20%" align="right"><font class="yazi">Etiket: </font></td>
		<td width="80%"><input type="text" name="etiket" class="alan" value="<%=etiket("etiket")%>" size="52"></td>
	</tr>
	<tr>
<%
set blog = Server.CreateObject("ADODB.Recordset")
SQL = "Select konu,id from blog where id = "&etiket("blog_id")&""
blog.open SQL,data,1,3
	blogisim=blog("konu")
blog.close : set blog=nothing
%>
		<td width="20%" align="right"><font class="yazi">Kayıtlı Blog: </font></td>
		<td width="80%">&nbsp;<font class="yazi"><%=blogisim%></font></td>
	</tr>
	<tr>
		<td width="20%"></td>
		<td width="80%"><input type="submit" class="dugme" value="Kaydet"></td>
	</tr>
</form>
</table>
</div>
<%
end if
etiket.close : set etiket = nothing
end if
end if
%>	
<%Else%>
</p>
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
<font class="baslik"><b>&nbsp;&nbsp;&nbsp;&nbsp;<%=ktg("konu")%></b></font>, <font class="yazi">Konusuna Ait</font>
<%
end if
ktg.close : set ktg = nothing
end if
end if
%>		
		Etiketler</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<%
siralama=filtre(Request.QueryString("siralama"))
tip=filtre(Request.QueryString("tip"))
sayi=filtre(Request.QueryString("sayi"))

	if siralama="1" then
		sira="blog_id"
	elseif siralama="2" then
		sira="etiket"
	elseif siralama="3" then
		sira="blog_id"
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
<table class="tablo" align="center" width="99%">
	<tr>
		<td width="22" height="24" class="ust"><font class="yazi">ID</font></td>
		<td width="786" height="24" class="ust"><font class="yazi">Etiket</font></td>
		<td width="155" height="24" class="ust"><font class="yazi">Bloga Git</font></td>
		<td width="177" height="24" class="ust"><font class="yazi">İşlem</font></td>
	</tr>
<%
if (Request.QueryString("Etiket"))="sil" then
id=request.querystring("id")
data.Execute("DELETE FROM etiket where id="&id&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if

Set zd_msg = Server.CreateObjecT("ADODB.recordSet")
if isnumeric(id)=false or id="" then
rSQL = "Select * from etiket order by "&sira&" "&t&""
else
rSQL = "Select * from etiket where blog_id="&id&" order by "&sira&" "&t&""
end if
zd_msg.open rSQL,data,1,3
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
		<td width="22" class="<%=stil%>"><font class="yazi"><%=zd_msg("id")%></td>
		<td width="786" class="<%=stil%>"><a href="etiket.asp?Etiketi=Duzenle&id=<%=zd_msg("id")%>"><%=zd_msg("etiket")%></a></td>
		<td width="155" class="<%=stil%>"><a href="../<%=SEOLink(zd_msg("id"))%>">GİT (<%=zd_msg("id")%>)</a></td>
		<td width="177" class="<%=stil%>">
<a href="?etiket=sil&id=<%=zd_msg("id")%>" onclick="return confirm('Silmek İstediğinizden Eminmisiniz?');">Sil</a> - <a href="../<%=SEOLink(zd_msg("blog_id"))%>" target="_blank">Bloga Git</a>
		</td>
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
<div align="center">
	<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="" method="get">
		<tr>
			<td>
		<input type="hidden" value="<%=id%>" name="id">
		<select name="siralama" size="1" class="alan">
        <option value="1"<%if sira="blog" then%> selected<%End if%>>Bloga Göre</option>
        <option value="2"<%if sira="etiket" then%> selected<%End if%>>İsme Göre</option>
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
<% End if %>
</body>

</html>
<% End if %>