<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Kategori Ekle</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<body background="images/arka.gif">
<%
if (Request.QueryString("kat"))="ekle" then

ad=request.form("ad")
aciklama=request.form("aciklama")

If ad="" or aciklama="" Then
Response.Write "<center>Boş Bıraktığınız Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history.back()""><<<GERİ</a></center>"
Response.End
End if

Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from kategori"
blgekle.Open SQL,data,1,3

blgekle.Addnew
blgekle("ad")=ad
blgekle("aciklama")=aciklama
blgekle.update

Response.Redirect("?kat=Eklendi")
End if

if (Request.QueryString("kat"))="Eklendi" then
response.write("<center>Kategori Eklenmiştir</center>")
End if
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Kategori Ekle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table width="99%" align="center" border"0" cellpadding="0" cellspacing="0" align="center" class="tablo">
	<form method="post" action="?kat=ekle">
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Ad :</font></div>
			</td>
			<td width="70%">
			<input name="ad" type="text" size="48" class="alan"></td>
		</tr>
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Açıklama : </font></div>
			</td>
			<td width="70%">
<textarea name="aciklama" class="alan" rows="6" cols="58"></textarea>
</td>

		</tr>
		<tr>
			<td width="30%"></td>
			<td width="70%">
			<input type="submit" value="Kaydet" class="dugme"></td>
		</tr>
	</form>
</table>

</body>

</html>
<% End if %>