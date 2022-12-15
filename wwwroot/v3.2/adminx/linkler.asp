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
<%
if (Request.QueryString("link"))="sil" then
id=request.querystring("id")
data.Execute("DELETE FROM linkler where id="&id&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")

elseif (Request.QueryString("link"))="duzenle" then
id=Request.QueryString("id")

if not isnumeric(id)=false or id="" then
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from linkler where id="&id&""
blgekle.open SQL,data,1,3
if not blgekle.eof then
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik"><%=blgekle("isim")%> Düzenleme</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
<table border="0" width="99%" id="table4" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="?link=kaydet&id=<%=id%>" method="post">
	<tr>
		<td width="30%" align="right"><font class="yazi">İsim :</font></td>
		<td width="70%"><input type="text" name="isim" class="alan" size="33" value="<%=blgekle("isim")%>"></td>
	</tr>
	<tr>
		<td width="30%" align="right"><font class="yazi">Link :</font></td>
		<td width="70%"><input type="text" name="link" class="alan" size="33" value="<%=blgekle("link")%>"></td>
	</tr>
	<tr>
		<td width="30%"></td>
		<td width="70%"><input type="submit" class="dugme" value="Kaydet"></td>
	</tr>
</form>
</table>
</div>
<%
end if
blgekle.close : set blgekle=nothing
end if
elseif (Request.QueryString("link"))="kaydet" then

isim=request.form("isim")
link=request.form("link")

If isim="" or link="" Then
Response.Write "<font class=yazi><center>Boş Bıraktığınız Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history.back()""><<<GERİ</a></center></font>"
Response.End
End if

id=Request.QueryString("id")
if not isnumeric(id)=false or id="" then
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from linkler where id="&id&""
blgekle.open SQL,data,1,3
if not blgekle.eof then

blgekle("isim")=isim
blgekle("link")=link

blgekle.update
response.redirect "linkler.asp"
end if
blgekle.close : set blgekle=nothing
end if

elseif (Request.QueryString("link"))="ekle" then
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Bağlantı Ekle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
<table border="0" width="99%" id="table4" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="?link=eklekaydet" method="post">
	<tr>
		<td width="30%" align="right"><font class="yazi">İsim :</font></td>
		<td width="70%"><input type="text" name="isim" class="alan" size="33"></td>
	</tr>
	<tr>
		<td width="30%" align="right"><font class="yazi">Link :</font></td>
		<td width="70%"><input type="text" name="link" class="alan" size="33" value="http://"></td>
	</tr>
	<tr>
		<td width="30%"></td>
		<td width="70%"><input type="submit" class="dugme" value="Kaydet"></td>
	</tr>
</form>
</table>
</div>
<%
elseif (Request.QueryString("link"))="eklekaydet" then

isim=request.form("isim")
link=request.form("link")

If isim="" or link="" Then
Response.Write "<font class=yazi><center>Boş Bıraktığınız Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history.back()""><<<GERİ</a></center></font>"
Response.End
End if

Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from linkler"
blgekle.Open SQL,data,1,3

blgekle.Addnew
blgekle("isim")=isim
blgekle("link")=link
blgekle.update

Response.Redirect "linkler.asp"
%>

<% Else %>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Bağlantılar Listeleniyor</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table border="0" width="99%" id="table2" style="border-collapse: collapse" class="tablo" align="center">
	<tr>
		<td height="24" class="ust" width="25"><font class="yazi">ID</font></td>
		<td height="24" class="ust" width="565"><font class="yazi">İsim</font></td>
		<td height="24" class="ust" width="418"><font class="yazi">Link</font></td>
		<td height="24" class="ust" width="213"><font class="yazi">İşlem</font></td>
	</tr>
<%
set zd_msg = Server.CreateObject("Adodb.Recordset")
SQL = "Select * from linkler order by id DESC"
zd_msg.open SQL,data,1,3
linksayisi="0"
mode = 2
Do While Not zd_msg.eof
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
linksayisi=linksayisi + 1
%>
	<tr>
		<td width="25" class="<%=stil%>" height="22"><font class="yazi"><%=zd_msg("id")%></font></td>
		<td width="565" class="<%=stil%>" height="22"><a href="?link=duzenle&id=<%=zd_msg("id")%>"><%=zd_msg("isim")%></a></td>
		<td width="418" class="<%=stil%>" height="22"><font class="yazi"><%=zd_msg("link")%></font></td>
		<td width="213" class="<%=stil%>" height="22"><a href="?link=sil&id=<%=zd_msg("id")%>" onclick="return confirm('Silmek İstediğinizden Eminmisiniz?');">Sil</a></td>
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
			<td align="center"><font class="yazi">Toplam Bağlantı : <%=linksayisi%></font></td>
		</tr>
	</table>
</div>
<% End if %>
</body>
</html>
<% end if %>