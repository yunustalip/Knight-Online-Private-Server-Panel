<!--#include file="ayar.asp"-->
<!--#include file="db.asp"-->
<!--#include file="inc.asp"-->
<!--#include file="filtre.asp"-->
<%
k=Request.QueryString("k")
b=Request.QueryString("b")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title><%=sitebaslik%> - Blog Arşivi</title>
<style>
a:link
{
font-size: 11px;
font-family: verdana;
color: #353535;
text-decoration: none;
border-bottom: 1px dotted gray;
display: inline;
}
a:visited
{
font-size: 11px;
font-family: verdana;
color: #353535;
text-decoration: none;
border-bottom: 1px dotted gray;
display: inline;
}
a:hover
{
font-size: 11px;
font-family: verdana;
color: #000000;
text-decoration: none;
border-bottom: 1px dotted red;
display: inline;
}
a:active
{
font-size: 11px;
font-family: verdana;
color: #353535;
text-decoration: none;
border-bottom: 1px dotted gray;
text-decoration: none;
}
.font { font-size: 11px; font-family: verdana; color: #353535; }
.tablo { border-collapse: collapse; border: 1px solid #C0C0C0; background-color: #FFFFFF; }
.bosluk { padding:5px; }
</style>
</head>

<body bgcolor="#EEEEEE">
<div align="center">
	<table border="0" width="100%" id="table1" cellpadding="0" class="tablo">
		<tr>
			<td>
			<br>
<div align="center">
	<table border="0" width="95%" id="table1" cellpadding="0" class="tablo" style="margin:3px;">
		<tr>
			<td class="bosluk"><a href="http://<%=strsite%>/"><span style="font-size:18px"><%=sitebaslik%> - Anasayfa</span></a></td>
		</tr>
		<tr>
			<td class="bosluk">&nbsp;&nbsp;- <a href="arsiv.asp"><span style="font-size:15px">Arşiv</span></a></td>
		</tr>
<%
if k<>"" then
if isnumeric(k) then
set kat=data.execute("select * from kategori where id="&k)
if not kat.eof then
kategoriad=kat("ad")
%>
		<tr>
			<td class="bosluk">&nbsp;&nbsp;&nbsp;&nbsp; - <a href="?k=<%=k%>"><%=kategoriad%></a>, <font class="font">(<%=kat("aciklama")%>)</font></td>
		</tr>
<%
end if
kat.close : set kat=nothing
end if
end if
%>
	</table>
</div>
<%
if isnumeric(k)=false or k="" then
set rs=data.execute("select * from kategori")
if not rs.eof then
%>
<div align="center">
	<table border="0" width="95%" id="table1" cellpadding="0" class="tablo" style="margin:3px;">
		<tr>
			<td class="bosluk"><font class="font" style="font-size:16px"><b>Kategoriler</b></font></td>
		</tr>
		<tr>
			<td class="bosluk">
<ol class="font">
<%
Do While Not rs.EOF

set blog=data.execute("select count(id) as blog_say from blog where kat_id= "&rs("id")&"")
%>
<li style="margin:3px"><a href="?k=<%=rs("id")%>"><%=rs("ad")%></a> <font class="font">[<%=blog(0)%>]</font></li>
<%
blog.close
set blog = Nothing

rs.MoveNext
Loop
%>
</ol>
			</td>
		</tr>
	</table>
</div>
<%
end if
rs.Close
Set rs = Nothing
end if
%>
<%
if b="" then
if k<>"" then
if isnumeric(k) then
set rs = data.execute("select id,konu,kat_id from blog where kat_id="&k&" order by id desc")
if not rs.eof then
%>
<div align="center">
	<table border="0" width="95%" id="table1" cellpadding="0" class="tablo" style="margin:3px;">
		<tr>
			<td class="bosluk"><font class="font" style="font-size:16px"><b><%=kategoriad%></b></font></td>
		</tr>
		<tr>
			<td class="bosluk">
<ol class="font">
<%
Do While Not rs.eof
set yorum=data.execute("select count(id) as yorum_say from yorum where onay=0 and blog_id="&rs("id"))
yorumsayisi=yorum(0)
yorum.close:set yorum=nothing
%>
<li style="margin:3px"><a href="?k=<%=rs("kat_id")%>&b=<%=rs("id")%>"><%=rs("konu")%></a> <font class="font">[<%=yorumsayisi%>]</font></li>
<% rs.movenext : loop %>
</ol>
			</td>
		</tr>
	</table>
</div>
<%
end if
rs.close : set rs=nothing
end if
end if
else
if isnumeric(b) and isnumeric(k) then
set rs = data.execute("select id,konu,mesaj,kat_id,hit,tarih from blog where id="&b&" and kat_id="&k&"")
if not rs.eof then
%>
<div align="center">
	<table border="0" width="95%" id="table1" cellpadding="0" class="tablo" style="margin:3px;">
		<tr>
			<td class="bosluk"><font class="font" style="font-size:16px"><b><%=rs("konu")%></b></font><br><font class="font" style="font-size:10px;color:gray"><%=rs("hit")%> defa okundu,
			</font></td>
		</tr>
		<tr>
			<td class="bosluk"><font class="font"><%=Replace(rs("mesaj"),"{KES}","")%></font></td>
		</tr>
	</table>
</div>
<%
set yorum=data.execute("select * from yorum where onay=0 and blog_id="&b&" order by tarih desc")
if not yorum.eof then
yorumsayisi=0
Do While Not yorum.eof
yorumsayisi=yorumsayisi + 1
%>
<div align="center">
	<table border="0" width="95%" id="table1" cellpadding="0" class="tablo" style="margin:3px;">
		<tr>
			<td bgcolor="#F9F9F9" height="24" class="bosluk"><font class="font"><b>Yazan:</b> <%=yorum("ekleyen")%>, 
			<b> <%=FormatDateTime(yorum("tarih"),1)%>&nbsp;<%=FormatDateTime(yorum("tarih"), 4)%></b></font></td>
		</tr>
		<tr>
			<td class="bosluk"><font class="font"><%=MesajFormatla(yorum("yorum"))%></font></td>
		</tr>
	</table>
</div>
<%
yorum.movenext : loop
%>
<center><font class="font"><b><%=yorumsayisi%> Yorum</b></font></center>
<%
end if
yorum.close : set yorum=nothing
end if
rs.close : set yorum=nothing
end if
end if
%>
			<br>
			</td>
		</tr>
	</table>
</div>
</body>

</html>
<%
data.close : set data = nothing
%>