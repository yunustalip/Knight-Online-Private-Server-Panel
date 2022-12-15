<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<body background="images/arka.gif">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Blog Kategorileri Listeleniyor</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table border="0" width="99%" id="table2" style="border-collapse: collapse" class="tablo" align="center">
	<tr>
		<td height="24" class="ust" width="22"><font class="yazi">ID</font></td>
		<td height="24" class="ust" width="875"><font class="yazi">İsim</font></td>
		<td height="24" class="ust" width="133"><font class="yazi">Bloglar</font></td>
		<td height="24" class="ust" width="189"><font class="yazi">İşlem</font></td>
	</tr>
	
<% response.buffer = "true" %>

<%
if (Request.QueryString("kat"))="sil" then
id=request.querystring("id")
data.Execute("DELETE FROM kategori where id like '"&id&"'")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if

set zd_msg = Server.CreateObject("Adodb.Recordset")
SQL = "Select * from kategori order by id DESC"
zd_msg.open SQL,data,1,3
kayitsayisi="0"
mode = 2
Do While Not zd_msg.eof
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
kayitsayisi=kayitsayisi + 1
%>
	<tr>
		<td width="22" class="<%=stil%>"><font class="yazi"><%=zd_msg("id")%></font></td>
		<td width="875" class="<%=stil%>"><a href="kategori_duz.asp?kat=duzenle&id=<%=zd_msg("id")%>"><%=zd_msg("ad")%></a></td>
		<td width="133" class="<%=stil%>"><%
set blog = Server.CreateObject("ADODB.RecordSet")
SQL = "select count(id) as blog_say from blog where kat_id= "&zd_msg("id")&""
blog.open SQL,data,1,3
%><a href="bloglar.asp?id=<%=zd_msg("id")%>">Blog: <b><%=blog("blog_say")%></b></a><%
blog.close
set blog = Nothing
%>		</td>
		<td width="189" class="<%=stil%>"><a href="?kat=sil&id=<%=zd_msg("id")%>" onclick="return confirm('Silmek İstediğinizden Eminmisiniz?');">Sil</a></td>
	</tr>
<%
zd_msg.movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
loop
%>
	</table>
<div align="center">
	<table border="0" width="99%" height="24" id="table3" cellpadding="0" style="border-collapse: collapse" class="tablo">
		<tr>
			<td align="center"><font class="yazi">Toplam Kategori : <%=kayitsayisi%></font></td>
		</tr>
	</table>
</div>
</body>
</html>
<% end if %>