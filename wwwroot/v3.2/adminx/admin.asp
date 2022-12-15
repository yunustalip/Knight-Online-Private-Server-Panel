<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Efendy Blog Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
</head>

<body background="images/arka.gif">
<%
adres=Request.ServerVariables("SCRIPT_NAME")
if not instr(adres,"/admin/")>0 then
%>
<div align="center">
	<table border="0" width="400" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td height="19" width="8">
			<img border="0" src="images/bas.gif" width="10" height="25"></td>
			<td height="19" width="1220" background="images/bg.gif">
			<p align="center"><font color="#6B816B"><b>
			<font face="Trebuchet MS" style="font-size: 14px">Admin</font><font face="Trebuchet MS" style="font-size: 14px"> 
			Paneli</font></b></font></td>
			<td height="19" width="11">
			<img border="0" src="images/son.gif" width="15" height="25"></td>
		</tr>
	</table>
</div>
<div align="center">
	<table border="0" width="394" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="kontrol.asp" method="post">
		<tr>
			<td width="149">&nbsp;</td>
			<td width="249">&nbsp;</td>
		</tr>
		<tr>
			<td width="149">
			<p align="right"><font class="yazi">
			Kullanıcı Adı :</font></td>
			<td width="249"><input type="text" name="kullanici" class="alan" style="width: 122"></td>
		</tr>
		<tr>
			<td width="149">
			<p align="right"><font class="yazi">
			Şifre :</font></td>
			<td width="249"><input type="password" name="sifre" class="alan" style="width: 122"></td>
		</tr>
		<tr>
			<td width="149" align="right"><font class="yazi">Bağlılık Süresi :</font></td>
			<td width="249">
<select name="sure" class="alan" style="width: 122">
<option value="10">10 dk</option>
<option value="30">30 dk</option>
<option value="60" selected>1 saat</option>
<option value="120">2 saat</option>
</select>
			</td>
		</tr>
		<tr>
			<td width="149">&nbsp;</td>
			<td width="249"><input type="submit" value="Giriş" class="dugme"></td>
		</tr>
</form>
	</table>
</div>
<%else%><center><font class="yazi"><b>Admin Paneli Klasörü Adını Değiştirmeden Panele Giremezsiniz. Bu bir güvenlik önlemidir</b></font></center><%end if%>
</body>

</html>