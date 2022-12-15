<!--#include file="db.asp"-->
<html>
<head>
<meta http-equiv="Content-Language" content="tr">
<%
kat=Request.QueryString("kat")
id=Request.QueryString("id")

if kat="" or isnumeric(kat)=false then
response.write "Hata 1"
else

set rgk = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri_kat where id="&kat&""
rgk.open SQL,data,1,3

if rgk.eof then
response.write "Hata 2"
else

set rg = Server.CreateObject("ADODB.RecordSet")
if id="" or isnumeric(id)=false then
SQL = "select * from galeri where kat_id="&rgk("id")&" order by id desc"
else
SQL = "select * from galeri where kat_id="&rgk("id")&" and id="&id&""
end if
rg.open SQL,data,1,3
if rg.eof then
response.write "Hata 3"
else
set kat = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as kat_say from galeri where kat_id="&rgk("id")&""
kat.open SQL,data,1,3
tumresim=kat("kat_say")
kat.close : set kat=nothing

set kat2 = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as kat_say from galeri where id >= "&rg("id")&" and kat_id="&rgk("id")&""
kat2.open SQL,data,1,3
if kat2("kat_say")="0" then
kalanresim="1"
else
kalanresim=kat2("kat_say")
end if
kat2.close : set kat2=nothing
%>
<title>Kategori: <%=rgk("isim")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<table id="Table_01" width="578" border="0" cellpadding="0" cellspacing="0" style="position: absolute; left: 0; top: 0">
	<tr>
		<td colspan="3" height="20" bgcolor="#DADADA"><font face="Tahoma">
		<font style="font-size: 11px; font-weight: 700;"><span style="float:right"><%=FormatDateTime(rg("tarih"),1)%>&nbsp;<%=FormatDateTime(rg("tarih"), 4)%></span><p></font></td>
	</tr>
	<tr>
		<td colspan="3" height="400"><img src="<%=rg("url")%>" width="578" height="400" border="0"></td>
	</tr>
	<tr>
		<td colspan="3" height="21" bgcolor="#DADADA"><center><font face="Tahoma" style="font-size: 11px"><%=rg("aciklama")%></font></center></td>
	</tr>
	<tr>
	<td height="33">
<table border="0" width="578" height="33" background="tema/images/slide/arkap.gif" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="43">
<%
Set Objonce = Server.CreateObject("adodb.RecordSet")
sql="SELECT id,kat_id From galeri WHERE id > "&rg("id")&" and kat_id="&rgk("id")&" order by id asc" 
Objonce.open sql, data, 1, 3
if not Objonce.eof then
%>
			<a href="?kat=<%=objonce("kat_id")%>&id=<%=objonce("id")%>"><img src="tema/images/slide/prew.gif" width="43" height="33" alt="" border="0"></a><%
end if
objonce.close : set objonce = nothing
%></td>
		<td width="490"><center><font face="Tahoma" style="font-size: 11px" color="#FFFFFF"><%=rg("isim")%>&nbsp;&nbsp;<%=kalanresim%>/<%=tumresim%></font></center></td>
		<td width="45">
<%
Set Objsonra = Server.CreateObject("adodb.RecordSet")
sql="SELECT id,kat_id From galeri WHERE id < "&rg("id")&" and kat_id="&rgk("id")&" order by id desc" 
Objsonra.open sql, data, 1, 3
if not Objsonra.eof then
%>
			<a href="?kat=<%=objsonra("kat_id")%>&id=<%=objsonra("id")%>"><img src="tema/images/slide/next.gif" width="45" height="33" alt="" border="0"></a><%
end if
Objsonra.close : set Objsonra = nothing 
%></td>
	</tr>
</table>
	</td>
	</tr>
</table>
<div style="position: relative; left: 8; top: 30"><a href="<%=rg("url")%>" target="_blank"><img border="0" src="tema/images/buyut.gif"></a></div>
<%
End if
rg.close : set rg=nothing
rgk.close : set rgk=nothing
End if
End if
%>
</body>
</html>