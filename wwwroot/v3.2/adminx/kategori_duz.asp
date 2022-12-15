<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; chablgekleet=windows-1254">
<title>Kategori Düzenle</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<body background="images/arka.gif">
<%
if (Request.QueryString("kat"))="kaydet" then

id=request.querystring("id")
ad=request.form("ad")
aciklama=request.form("aciklama")

If ad="" or aciklama="" Then
Response.Write "<font class=""yazi"">Boþ Býraktýðýnýz Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history.back()""><<<GER&#304;</a></font>"
Response.End
End if

set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from kategori where id="&id&""
blgekle.Open SQL,data,1,3
if not blgekle.eof then
blgekle("ad")=ad
blgekle("aciklama")=aciklama

blgekle.update
blgekle.Close
end if
Set blgekle = Nothing


Response.Redirect("kategoriler.asp")
End if

if (Request.QueryString("kat"))="duzenle" then

id=request.querystring("id")
if not isnumeric(id)=false then
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from kategori where id="&id&""
blgekle.open SQL,data,1,3
if not blgekle.eof then
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Kategori Düzenle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table width="99%" align="center" border"0" cellpadding="0" cellspacing="0" align="center" class="tablo">
	<form method="post" action="?kat=kaydet&id=<%=id%>">
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Ad :</font></div>
			</td>
			<td width="70%">
			<input name="ad" type="text" size="48" class="alan" value="<%=blgekle("ad")%>"></td>
		</tr>
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Açýklama : </font></div>
			</td>
			<td width="70%">
<textarea name="aciklama" class="alan" rows="6" cols="58"><%=blgekle("aciklama")%></textarea>
</td>

		</tr>
		<tr>
			<td width="30%"></td>
			<td width="70%">
			<input type="submit" value="Kaydet" class="dugme"></td>
		</tr>
	</form>
</table><%
End if
blgekle.Close
Set blgekle = Nothing
Response.End
End if
End if
%>
</body>

</html>
<% end if %>