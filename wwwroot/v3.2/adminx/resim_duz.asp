<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; chablgekleet=windows-1254">
<title>Resim Düzenle</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<!--#include file="db.asp"-->
<%
if (Request.QueryString("resim"))="kaydet" then

id=request.querystring("id")
if not isnumeric(id)=false then
isim=request.form("isim")
aciklama=request.form("aciklama")
url=request.form("url")
url_kucuk=request.form("url_kucuk")
kat_id=request.form("kat_id")


If isim="" or url="" or url_kucuk="" or kat_id="" Then
Response.Write "<center><br><br><center>Bo&#351; B&#305;rakt&#305;&#287;&#305;n&#305;z Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history.back()""><<<GER&#304;</a></center></center>"
Response.End
End if

set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri where id="&id&""
blgekle.Open SQL,data,1,3
if not blgekle.eof then
blgekle("aciklama")=aciklama
blgekle("isim")=isim
blgekle("url_kucuk")=url_kucuk
blgekle("url")=url
blgekle("kat_id")=kat_id

blgekle.update
end if
blgekle.Close
Set blgekle = Nothing


Response.Redirect("resimler.asp")
end if
End if

if (Request.QueryString("resim"))="duzenle" then

id=request.querystring("id")
	if not isnumeric(id)=false then
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from galeri where id="&id&""
blgekle.open SQL,data,1,3
if not blgekle.eof then
%>
<body background="images/arka.gif">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Resim Düzenle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table width="99%" align="center" border="0" cellpadding="0" cellspacing="0" class="tablo" align="center">
	<form method="post" action="?resim=kaydet&id=<%=id%>" name="rg">
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Ýsim :</font></div>
			</td>
			<td width="70%">
			<input name="isim" type="text" size="34" class="alan" value="<%=blgekle("isim")%>"></td>
		</tr>
		<tr>
			<td width="30%" align="right">
			<font class="yazi">Büyük Resim Adresi:</font></td>
			<td width="70%">
			<input name="url" type="text" size="34" class="alan" value="<%=blgekle("url")%>"><a ONCLICK="window.open('yukle.asp?yazdir=url','resimyukle','top=20,left=20,width=450,height=300,toolbar=no,scrollbars=yes');" href="#">Bilgisayardan Yükle</a></td>
		</tr>
		<tr>
			<td width="30%" align="right">
			<font class="yazi">Küçük Resim Adresi:</font></td>
			<td width="70%">
			<input name="url_kucuk" type="text" size="34" class="alan" value="<%=blgekle("url_kucuk")%>"><a ONCLICK="window.open('yukle.asp?yazdir=url_kucuk','resimyukle','top=20,left=20,width=450,height=300,toolbar=no,scrollbars=yes');" href="#">Bilgisayardan Yükle</a></td>
		</tr>
		<tr>
			<td width="30%">
			<div align="right">
			<font class="yazi">Açýklama : </font></div>
			</td>
			<td width="70%">
<textarea name="aciklama" class="alan" rows="6" cols="67"><%=blgekle("aciklama")%></textarea>
			</td>

		</tr>
		<tr>
			<td width="30%" align="right">
			<font class="yazi">Kategori :</font></td>
			<td width="70%">
<select name="kat_id" class="alan">
<%
set kate = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from galeri_kat"
kate.open SQL,data,1,3
do while not kate.eof
%>
							<option value="<% = kate("id") %>"<%if id=kate("id") then%> selected<%end if%>><% = kate("isim") %></option>
<%
kate.movenext : loop
kate.close : set kate=nothing
%>
			</td>
		</tr>
		<tr>
			<td width="30%"></td>
			<td width="70%">
			<input type="submit" value="Kaydet" class="dugme"></td>
		</tr>
	</form>
</table>
<%
blgekle.Close
Set blgekle = Nothing
Response.End
End if
End if
End if
%>
</body>

</html>
<% end if %>