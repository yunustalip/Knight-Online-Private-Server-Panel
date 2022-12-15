<% if session("admin") then %>
<!--#include file="db.asp"-->
<!--#include file="../filtre.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Admin Paneli</title>
</head>
<style>
<!--
BODY 
{
font: 9px trebuchet ms;
color: #000000;
scrollbar-arrow-color:#000000;
scrollbar-track-color:#ECE9D8;
scrollbar-shadow-color:#ACA899;
scrollbar-face-color:#ECE9D8;
scrollbar-highlight-color:#ECE9D8;
scrollbar-darkshadow-color:#ECE9D8;
scrollbar-3dlight-color:#ECE9D8;
}
.mouse
{
cursor: hand;
font-family: trebuchet ms;
color: #000000;
font: 12px trebuchet ms;
color: #000000;
}
-->
</style>
<body bgcolor="#ECE9D8" topmargin="4" leftmargin="4" rightmargin="4" bottommargin="4">
<table border="1" width="100%" id="table1" class="mouse" cellspacing="0" cellpadding="0" bordercolor="#CCC8B8" style="border-collapse: collapse" height="72">
	<tr>
		<td>
		<p align="right"><b><font size="3">EFENDY BLOG ADMIN PANELINE HOŞGELDİNİZ</font></b><font size="3"><br>
		</font><font style="font-size: 11px">
		<br>
		Öncelikle Sol Menüden Yapacağınız İşlemi Seçin</font></td>
	</tr>
</table>
<font style="font-size: 11px">
<br>
</font>
<table border="1" width="100%" id="table2" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#ECE9D8">
	<tr>
		<td width="50%" bgcolor="#CCC8B8">
		<p align="center"><font style="font-size: 11px">Son Eklenen Bloglar</font></td>
		<td bgcolor="#CCC8B8">
		<p align="center"><font style="font-size: 11px">Son Eklenen Resimler</font></td>
	</tr>
	<tr>
		<td width="50%">
		<table border="1" width="100%" id="table3" cellpadding="0" style="border-collapse: collapse" height="100%" bordercolor="#CCC8B8">
<%
set blog = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from blog order by id DESC"
blog.open SQL,data,1,3

For p = 1 To 5
if blog.eof Then exit For
%>
			<tr>
				<td width="364"><font style="font-size: 11px"><%=blog("konu")%></font></td>
				<td width="51">
				<p align="center"><a href="blog_duz.asp?Blog=duzenle&id=<%=blog("id")%>">
				<span style="text-decoration: none">
				<font color="#000000" style="font-size: 11px">Düzenle</font></span></a></td>
			</tr>
<% 
blog.Movenext 
Next
blog.Close
Set blog = Nothing
%>
		</table>
		</td>
		<td>
		<table border="1" width="100%" id="table4" cellpadding="0" style="border-collapse: collapse" bordercolor="#CCC8B8">
<%
set resim = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri order by tarih DESC"
resim.open SQL,data,1,3

For p = 1 To 5
if resim.eof Then exit For
%>
			<tr>
				<td width="435"><font style="font-size: 11px"><%=resim("isim")%></font></td>
				<td width="62">
				<p align="center"><a href="resim_duz.asp?resim=duzenle&id=<%=resim("id")%>">
				<span style="text-decoration: none">
				<font color="#000000" style="font-size: 11px">Düzenle</font></span></a></td>
			</tr>
<% 
resim.Movenext 
Next
resim.Close
Set resim = Nothing
%>
		</table>
		</td>
	</tr>
	<tr>
		<td width="50%" bgcolor="#CCC8B8">
		<p align="center"><font style="font-size: 11px">Son Yapılan Yorumlar</font></td>
	</tr>
	<tr>
		<td width="50%">
		<table border="1" width="100%" id="table5" cellpadding="0" style="border-collapse: collapse" height="100%" bordercolor="#CCC8B8">
<%
set yorum = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from yorum where onay=0 order by tarih DESC"
yorum.open SQL,data,1,3

For p = 1 To 5
if yorum.eof Then exit For
%>
			<tr>
				<td width="365"><font style="font-size: 11px"><%=Left(yorum("yorum"),50)%></font></td>
				<td width="51">
<font style="font-size: 11px">
<%
set ktg = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from blog where id = "&yorum("blog_id") &""
ktg.open SQL,data,1,3
if ktg.eof then
else
end if
if yorum("blog_id")="0" then
link="hakkimda.asp#yorumlar"
else
link="../"&SEOLink(yorum("blog_id"))
end if
%>
				</font>
				<p align="center"><font style="font-size: 11px"><a href="../<%=link%>">
				<span style="text-decoration: none"><font color="#000000"><%if yorum("blog_id")="0" then%>HAKKIMDA<%else%>GİT (<%=yorum("blog_id")%>)<%end if%></font></span></a>
<% ktg.close : set ktg = nothing %>
				</font>
				</td>
			</tr>
<% 
yorum.Movenext 
Next
yorum.Close
Set yorum = Nothing
%>
		</table>
		</td>
		<td valign="top">
		</td>
	</tr>
</table>
<font style="font-size: 11px">
<br>
</font>
<table border="1" width="100%" id="table7" style="border-collapse: collapse" bordercolor="#ECE9D8">
	<tr>
		<td width="100%" colspan="2" height="24" bgcolor="#CCC8B8" bordercolor="#000000" style="border: 1px solid #000000">
		<p align="center"><b><font style="font-size: 11px">İstatistikler</font></b></td>
	</tr>
	<tr>
		<td width="50%">
		<table border="1" width="100%" id="table8" cellpadding="0" style="border-collapse: collapse" bordercolor="#CCC8B8">
			<tr>
<%
set kat = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as kat_say from kategori"
kat.open SQL,data,1,3
%>
				<td width="121"><font style="font-size: 11px">Toplam Kategori</font></td>
				<td width="367"><font style="font-size: 11px"><%=kat("kat_say")%></font></td>
<%
kat.close
set kat = Nothing
%>
			</tr>
			<tr>
<%
set blg = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as blg_say from blog"
blg.open SQL,data,1,3
%>
				<td width="121"><font style="font-size: 11px">Toplam Blog</font></td>
				<td width="367"><font style="font-size: 11px"><%=blg("blg_say")%></font></td>
<%
blg.close
set blg = Nothing
%>
			</tr>
			<tr>
<%
set yrm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yrm_say from yorum"
yrm.open SQL,data,1,3
%>
				<td width="121"><font style="font-size: 11px">Toplam Yorum</font></td>
				<td width="367"><font style="font-size: 11px"><%=yrm("yrm_say")%></font></td>
<%
yrm.close
set yrm = Nothing
%>
			</tr>
			<tr>
<%
set tm_o = Server.CreateObject("ADODB.RecordSet")
SQL = "select SUM(hit) as tm_o from blog"
tm_o.open SQL,data,1,3
%>
				<td width="121"><font style="font-size: 11px">Toplam Okunma</font></td>
				<td width="367"><font style="font-size: 11px"><%=tm_o("tm_o")%></font></td>
<%
tm_o.close
set tm_o = Nothing
%>
			</tr>
		</table>
		</td>
		<td width="50%" valign="top">
		<table border="1" width="100%" id="table9" cellpadding="0" style="border-collapse: collapse" bordercolor="#CCC8B8">
			<tr>
<%
set rsm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as rsm_say from galeri"
rsm.open SQL,data,1,3
%>
				<td width="152"><font style="font-size: 11px">Toplam Resim</font></td>
				<td width="339"><font style="font-size: 11px"><%=rsm("rsm_say")%></font></td>
<%
rsm.close
set rsm = Nothing
%>
			</tr>
			<tr>
<%
set zd = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as zd_say from zd"
zd.open SQL,data,1,3
%>
				<td width="152"><font style="font-size: 11px">Toplam Mesaj (zd)</font></td>
				<td width="339"><font style="font-size: 11px"><%=zd("zd_say")%></font></td>
<%
zd.close
set zd = Nothing
%>
			</tr>
			<tr>
<%
set ileti = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as ileti_say from iletisim"
ileti.open SQL,data,1,3
%>
				<td width="152"><font style="font-size: 11px">Toplam İleti</font></td>
				<td width="339"><font style="font-size: 11px"><%=ileti("ileti_say")%></font></td>
<%
ileti.close
set ileti = Nothing
%>
			</tr>
		</table>
		</td>
	</tr>
</table>
<font style="font-size: 11px">
<br>
</font>
<table border="1" width="100%" id="table10" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#CCC8B8">
	<tr>
		<td style="border-style: solid; border-width: 1px; background-color: #CCC8B8" bordercolor="#000000" height="25" width="995" colspan="2">
		<p align="center"><b><font style="font-size: 11px">Onay Bekleyenler</font></b></td>
	</tr>
	<tr>
<%
set yor = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as yor from yorum where onay=1"
yor.open SQL,data,1,3
%>
		<td width="50%"><font style="font-size: 11px">Yorumlar: <%=yor("yor")%></font></td>
<%
yor.close
set yor = Nothing

set zdm = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as zdm from zd where onay=1"
zdm.open SQL,data,1,3
%>
		<td width="50%"><font style="font-size: 11px">Mesajlar(zd): <%=zdm("zdm")%></font></td>
<%
zdm.close
set zdm = Nothing
%>
	</tr>
</table>
<% end if %>