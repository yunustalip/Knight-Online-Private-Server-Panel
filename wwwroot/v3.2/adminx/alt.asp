<% if session("admin") then %>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Efendy Blog Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
</head>

<body background="images/arka.gif">
<!--#include file="db.asp"-->
<!--#include file="../filtre.asp"-->
<%
set kat = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as kat_say from kategori"
kat.open SQL,data,1,3
	kategorisayisi=kat("kat_say")
kat.close
set kat = Nothing

set blg = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as blg_say from blog"
blg.open SQL,data,1,3
	blogsayisi=blg("blg_say")
blg.close
set blg = Nothing

set yrm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yrm_say from yorum"
yrm.open SQL,data,1,3
	yorumsayisi=yrm("yrm_say")
yrm.close
set yrm = Nothing

set tm_o = Server.CreateObject("ADODB.RecordSet")
SQL = "select SUM(hit) as tm_o from blog"
tm_o.open SQL,data,1,3
	toplamokunma=tm_o("tm_o")
tm_o.close
set tm_o = Nothing

set rsm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as rsm_say from galeri"
rsm.open SQL,data,1,3
	resimsayisi=rsm("rsm_say")
rsm.close
set rsm = Nothing

set zd = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as zd_say from zd"
zd.open SQL,data,1,3
	mesajsayisi=zd("zd_say")
zd.close
set zd = Nothing

set ileti = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as ileti_say from iletisim"
ileti.open SQL,data,1,3
	iletisayisi=ileti("ileti_say")
ileti.close
set ileti = Nothing

set yor = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yor from yorum where onay=1"
yor.open SQL,data,1,3
	onaybekleyenyorum=yor("yor")
yor.close
set yor = Nothing

set zdm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as zdm from zd where onay=1"
zdm.open SQL,data,1,3
	onaybekleyenmesaj=zdm("zdm")
zdm.close
set zdm = Nothing
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td align="left" width="50%">

		<table border="0" width="99.5%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
			<tr>
		<td height="25" class="ust" colspan="2">
		<p align="center"><font class="yazi">Son Eklenen Bloglar</font></p></td>
			</tr>
<%
set blog = Server.CreateObject("ADODB.RecordSet")
SQL = "select konu,id from blog order by id DESC"
blog.open SQL,data,1,3
mode = 2
For p = 1 To 5
if blog.eof Then exit For
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
%>
			<tr>
				<td width="90%" class="<%=stil%>" height="20"><font class="yazi"><%=blog("konu")%></font></td>
				<td width="10%" class="<%=stil%>" height="20">
				<p align="center"><a href="blog_duz.asp?Blog=duzenle&id=<%=blog("id")%>">Düzenle</a></td>
			</tr>
<% 
blog.Movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
Next
blog.Close
Set blog = Nothing
%>
		</table>

		</td>
		<td align="right" width="50%">
		
		<table border="0" width="99.5%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
			<tr>
			<td height="25" class="ust" colspan="2"><p align="center"><font class="yazi">Son Eklenen Resimler</font></p></td>
			</tr>
<%
set resim = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri order by tarih DESC"
resim.open SQL,data,1,3
mode = 2
For p = 1 To 5
if resim.eof Then exit For
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
%>
			<tr>
				<td width="90%" class="<%=stil%>" height="20"><font class="yazi"><%=resim("isim")%></font></td>
				<td width="10%" class="<%=stil%>" height="20">
				<p align="center"><a href="resim_duz.asp?resim=duzenle&id=<%=resim("id")%>">Düzenle</a></td>
			</tr>
<% 
resim.Movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
Next
resim.Close
Set resim = Nothing
%>
		</table>
		
		</td>
	</tr>
	<tr>
		<td align="left" width="50%">
		
		<table border="0" width="99.5%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
	<tr>
		<td width="50%" height="25" class="ust" colspan="2">
		<p align="center"><font class="yazi">Son Yorumlar</font></p></td>
		<td width="50%" height="25"></td>
	</tr>
<%
set yorum = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from yorum where onay=0 order by tarih DESC"
yorum.open SQL,data,1,3
mode = 2
For p = 1 To 5
if yorum.eof Then exit For
set ktg = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from blog where id = "&yorum("blog_id") &""
ktg.open SQL,data,1,3
if ktg.eof then
else
end if
if yorum("blog_id")="0" then
link="../hakkimda.asp#yorumlar"
else
link="../"&SEOLink(yorum("blog_id"))
end if
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
%>
			<tr>
				<td width="90%" class="<%=stil%>" height="20"><font class="yazi"><%=Left(Cevir(yorum("yorum")),65)%></font></td>
				<td width="10%" class="<%=stil%>" height="20"><a href="../<%=link%>">
				<p align="center"><%if yorum("blog_id")="0" then%>HAKKIMDA<%else%>GİT (<%=yorum("blog_id")%>)<%end if%></a></td>
			</tr>
<%
ktg.close : set ktg = nothing
yorum.Movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
Next
yorum.Close
Set yorum = Nothing
%>
		</table>
		
		</td>
		<td valign="top" align="right" width="50%">
<table border="0" width="99.5%" id="table1" cellpadding="0" style="border-collapse: collapse;" class="tablo">
	<tr>
		<td width="100%" colspan="2" align="center" height="25" class="ust"><font class="yazi">Onay Bekleyenler</font></td>
	</tr>
	<tr>
		<td height="20"><a href="yorum_onay.asp">Yorum: <%=onaybekleyenyorum%></a></td>
	</tr>
	<tr>
		<td height="20" class="tablo1"><a href="zd_onay.asp">Mesaj(zd): <%=onaybekleyenmesaj%></a></td>
	</tr>
</table>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse" class="tablo">
	<tr>
		<td width="100%" colspan="2" align="center" height="25" class="ust"><font class="yazi">Site İstatistikleri</font></td>
	</tr>
	<tr>
		<td width="50%" class="tablo1" height="20"><font class="yazi">Toplam Kategori: <%=kategorisayisi%></font></td>
		<td width="50%" class="tablo1" height="20"><font class="yazi">Toplam Resim: <%=resimsayisi%></font></td>
	</tr>
	<tr>
		<td width="50%" height="20"><font class="yazi">Toplam Blog: <%=blogsayisi%></font></td>
		<td width="50%" height="20"><font class="yazi">Toplam Mesaj(zd): <%=mesajsayisi%></font></td>
	</tr>
	<tr>
		<td width="50%" class="tablo1" height="20"><font class="yazi">Toplam Yorum: <%=yorumsayisi%></font></td>
		<td width="50%" class="tablo1" height="20"><a href="ileti.asp">Toplam İleti: <%=iletisayisi%></a></td>
	</tr>
	<tr>
		<td width="50%" height="20"><font class="yazi">Toplam Okunma(blog): <%=toplamokunma%></font></td>
		<td width="50%" height="20"><font class="yazi"></font></td>
	</tr>
</table>
		</td>
	</tr>
</table>
</body>

</html>
<% End if %>