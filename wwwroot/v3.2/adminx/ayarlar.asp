<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Site Ayarları</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="../ayar.asp"-->
<!--#include file="db.asp"-->
<!--#include file="../inc.asp"-->
<%
if (Request.QueryString("ayarlari"))="kaydet" then

Set ayar = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from ayar"
ayar.Open SQL,data,1,3

ayar("sitebaslik")=request.form("sitebaslik")
ayar("site")=request.form("site")
ayar("adkull")=request.form("adkull")
ayar("adsif")=request.form("adsif")
ayar("aciklama")=request.form("aciklama")
ayar("etiket")=request.form("etiket")
ayar.update

ayar.Close
Set ayar = Nothing

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if
%>
<body background="images/arka.gif">

<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Site Ayarları</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table class="tablo" width="99%" align="center">
<form action="?ayarlari=kaydet" method="post">
	<tr>
		<td width="30%" height="20" align="right"><font class="yazi">Site Başlık :</font></td>
		<td width="70%" height="20"><input name="sitebaslik" size="41" value="<%=sitebaslik%>" class="alan"></td>
	</tr>
	<tr>
		<td width="30%"></td>
		<td width="70%"><font class="yazi" style="font-weight:normal">Sitenin <u>title</u> kısmında görünecektir</font></td>
	</tr>
	<tr>
		<td width="30%" height="20" align="right"><font class="yazi">Site Adresi :</font></td>
		<td width="70%" height="20"><input name="site" size="41" value="<%=strsite%>" class="alan"></td>
	</tr>
	<tr>
		<td width="30%"></td>
		<td width="70%"><font class="yazi" style="font-weight:normal"><u>http://</u> kullanmayınız adresin sonuna <u>/</u> eklemeyin.</font></td>
	</tr>
	<tr>
		<td width="30%" height="20" align="right"><font class="yazi">Siteyle İlgili Açıklama :</font></td>
		<td width="70%" height="20"><input name="aciklama" size="41" value="<%=aciklama%>" class="alan"></td>
	</tr>
	<tr>
		<td width="30%"></td>
		<td width="70%"><font class="yazi" style="font-weight:normal">Siteyi Açıklayan Bir Cümle</font></td>
	</tr>
	<tr>
		<td width="30%" height="20" align="right"><font class="yazi">Etiketler :</font></td>
		<td width="70%" height="20"><input name="etiket" size="41" value="<%=etiket%>" class="alan"></td>
	</tr>
	<tr>
		<td width="30%"></td>
		<td width="70%"><font class="yazi" style="font-weight:normal">Site ilgili etiketler <u>virgül(,)</u> ile ayırın.</font></td>
	</tr>
	<tr>
		<td width="30%" class="tablo1" height="20" align="right"><font class="yazi">Admin Kullanıcı Adı :</font></td>
		<td width="70%" class="tablo1" height="20"><input name="adkull" size="41" value="<%=adkull%>" class="alan"></td>
	</tr>
	<tr>
		<td width="30%" class="tablo1" height="20" align="right"><font class="yazi">Admin Parolası :</font></td>
		<td width="70%" class="tablo1" height="20"><input type="password" name="adsif" size="41" value="<%=adsif%>" class="alan"></td>
	</tr>
	<tr>
		<td width="30%" height="20" align="right"></td>
		<td width="70%" height="20"><input type="Submit" value="Kaydet" class="dugme"></td>
	</tr>
</form>
</table>
</body>

</html>
<% End if %>