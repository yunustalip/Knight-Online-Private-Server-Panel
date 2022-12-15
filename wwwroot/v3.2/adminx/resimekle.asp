<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Resim Ekle</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<%
if (Request.QueryString("resim"))="ekle" then

isim=request.form("isim")
aciklama=request.form("aciklama")
kat_id=request.form("kat_id")
url=request.form("url")
url_kucuk=request.form("url_kucuk")

If isim="" or aciklama="" or url="" or url_kucuk="" or kat_id="" Then
Response.Write "<center>Boş Bıraktığınız Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history.back()""><<<GERİ</a></center>"
Response.End
End if

Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from galeri"
blgekle.Open SQL,data,1,3

blgekle.Addnew
blgekle("aciklama")=aciklama
blgekle("url")=url
blgekle("url_kucuk")=url_kucuk
blgekle("isim")=isim
blgekle("kat_id")=kat_id
blgekle("Tarih")=now()
blgekle.update

Response.Redirect("?resim=Eklendi")
End if

if (Request.QueryString("resim"))="Eklendi" then
response.write("<center>Resim Eklenmiştir</center>")
End if
%>
<body background="images/arka.gif">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Resim Ekle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table width="99%" align="center" border="0" cellpadding="0" cellspacing="0" class="tablo" align="center">
	<form method="post" action="?resim=ekle" name="rg">
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">İsim :</font></div>
			</td>
			<td width="70%">
			<input name="isim" type="text" size="34" class="alan"></td>
		</tr>
		<tr>
			<td width="30%">
			<p align="right"><font class="yazi">Büyük Resim Adresi:</font></td>
			<td width="70%">
			<input name="url" type="text" size="34" class="alan"><a ONCLICK="window.open('yukle.asp?yazdir=url','resimyukle','top=20,left=20,width=450,height=300,toolbar=no,scrollbars=yes');" href="#">Bilgisayardan Yükle</a></td>
		</tr>
		<tr>
			<td width="30%">
			<p align="right"><font class="yazi">Küçük Resim Adresi:</font></td>
			<td width="70%">
			<input name="url_kucuk" type="text" size="34" class="alan"><a ONCLICK="window.open('yukle.asp?yazdir=url_kucuk','resimyukle','top=20,left=20,width=450,height=300,toolbar=no,scrollbars=yes');" href="#">Bilgisayardan Yükle</a></td>
		</tr>
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Açıklama : </font></div>
			</td>
			<td width="70%">
<textarea name="aciklama" class="alan" rows="6" cols="67"></textarea>

</td>

		</tr>
		<tr>
			<td width="30%" height="24">
			<p align="right"><font class="yazi">Kategori :</font></td>
			<td width="70%">
<select name="kat_id" class="alan">
<option selected value="">Kategori seçin</option>
<%
set katsec = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri_kat"
katsec.open SQL,data,1,3
Do While Not katsec.EOF
%>
<option value="<%=katsec("id")%>"><%=katsec("isim")%></option>
<%
katsec.MoveNext
Loop
katsec.Close
Set katsec = Nothing
%>
</select>
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