<% if session("admin") then %>
<!--#include file="db.asp"-->
<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
</head>
<%
if (request.querystring("sik"))="sil" then
id=request.querystring("id")
	data.Execute("DELETE FROM anket where id="&id&"")
end if
if (request.querystring("anket"))="sil" then
id=request.querystring("id")
data.Execute("DELETE FROM ankets where id like '"&id&"'")
data.Execute("DELETE FROM anket where a_id like '"&id&"'")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if

if (request.querystring("cevap"))="duzenle" then
id=request.querystring("id")
	cevap=Request.Form("cevap")

set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from anket where id="&id&""
blgekle.Open SQL,data,1,3

blgekle("cevap")=cevap

blgekle.update
blgekle.Close : set blgekle=nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
end if

if (request.querystring("sik"))="ekle" then
a_id=request.querystring("a_id")
	cevap=Request.Form("cevap")
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from anket"
blgekle.Open SQL,data,1,3
blgekle.addnew
blgekle("cevap")=cevap
blgekle("a_id")=a_id
blgekle("deger")=0
blgekle.update
blgekle.Close : set blgekle=nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
end if


if (request.querystring("anket"))="duzenle" then
id=request.querystring("id")
	soru=Request.Form("soru")

set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from ankets where id="&id&""
blgekle.Open SQL,data,1,3

blgekle("soru")=soru

blgekle.update
blgekle.Close : set blgekle=nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
end if

if (Request.QueryString("anket"))="ekle" Then
soru=request.form("soru")
If soru="" Then
Response.Write "baslik giriniz."
Response.End
Else
data.Execute ("INSERT INTO ankets (soru) VALUES  ('"&soru&"')")
Set soru = data.Execute("Select id from ankets Order by id desc") 
For i = 1 To request.form("cevap").Count
scenekveri=request.form("cevap")(i)
if not scenekveri="" then
scenekveri = Replace(scenekveri, Chr(39), "&#39;", 1, -1, 1)
data.Execute ("INSERT INTO anket (cevap,a_id,deger) VALUES  ('"&scenekveri&"' , '"&soru("id")&"' , '0')")
end if
Next
soru.Close
Set soru = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if
end if
%>
<body background="images/arka.gif">
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Anket Ekle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table border="0" width="99%" id="table1" cellpadding="0" style="border-collapse: collapse" class="tablo" align="center">
<form method="POST" action="">
	<tr>
		<td valign="top"><font class="yazi">Anket İçin Şık 
		Sayısını Girin:</font>
			<input type="text" name="sayi" size="33" class="alan"><input type="submit" value="Tamam" class="dugme">
		</td>
	</tr>
</form>
	<tr>
		<td valign="top">
		<table border="0" width="100%" id="table2" cellpadding="0" style="border-collapse: collapse">
<%
sayi=Request.Form("sayi")
If sayi="" or isnumeric(sayi)=false Then
sayi = 5
End if
%>
<form action="?anket=ekle" method="post">
			<tr>
				<td width="30%" class="tablo1">
				<p align="right"><font class="yazi">Anket İçin Başlık 
				:</font></td>
				<td width="70%" class="tablo1">
				<input type="text" name="soru" size="33" class="alan"></td>
			</tr>
<%
a=0
For i=1 To sayi
a = a + 1
%>
			<tr>
				<td width="30%">
				<p align="right"><font  class="yazi">Şık 
				<%=a%> :</font></p></td>
				<td width="70%"><font  class="yazi">
				<input type="text" name="cevap" size="33" class="alan"></font></td>
			</tr>
<% next %>
			<tr>
				<td width="30%">&nbsp;</td>
				<td width="70%"><font class="yazi">
				<input type="submit" value="Tamam" class="dugme"></font></td>
			</tr>
</form>
		</table>
		</td>
	</tr>
</table>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Anket Yönetimi</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table border="0" width="99%" id="table3" cellpadding="0" style="border-collapse: collapse" class="tablo" align="center">
<%
set rs = data.execute("SELECT * FROM ankets order by id desc")
do while not rs.eof 
%>
		<form action='?anket=duzenle&id=<%=rs("id")%>' method='post'>
	<tr>
		<td width="95%" class="ust">
		<input type="text" value='<%=rs("soru")%>' name='soru' class="alan" size="33"><input type="submit" value="Kaydet" class="dugme"></td>
		<td width="5%" class="ust">
		<a href="?anket=sil&id=<%=rs("id")%>" onclick="return confirm('Silmek İstediğinizden Eminmisiniz?');">Sil</a></td>
	</tr>
		</form>

<%
set rx = data.execute("SELECT * FROM anket where a_id="&rs("id")&"")
do while not rx.eof
%>
		<form action="?cevap=duzenle&id=<%=rx("id")%>" method="post">
	<tr>
		<td colspan="2"><span style="float:left">
		<input type="text" value="<%=rx("cevap")%>" name="cevap" size="33" class="alan"><input type="submit" value="Kaydet" class="dugme"><font class="yazi">OY: <%=rx("deger")%></font></span><span style="float:right">
<a href="?sik=sil&id=<%=rx("id")%>" onclick="return confirm('Silmek İstediğinizden Eminmisiniz?');">Sil</a>
		&nbsp;</td>
	</tr>
		</form>
<%
rx.movenext : loop : rx.close : set rx=nothing
%>
	<form action='?sik=ekle&a_id=<%=rs("id")%>' method='post'>
	<tr>
		<td colspan='2'><input name='cevap' type='text' class="alan" size="33"><input type='submit' value='Şık Ekle' class="dugme"></td>
	</tr>
	</form>
	<tr>
		<td colspan="2" class="tablo1" height="20"></td>
	</tr>
<%
rs.movenext
loop
rs.close
set rs = nothing
%>
</table>
<% End if %>