<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; chablgekleet=windows-1254">
<title>Blog Ekle</title>
		<script type="text/javascript" src="scripts/wysiwyg.js"></script>
		<script type="text/javascript" src="scripts/wysiwyg-settings.js"></script>
		<!-- 
			Attach the editor on the textareas
		-->
		<script type="text/javascript">
			// Use it to attach the editor to all textareas with full featured setup
			//WYSIWYG.attach('all', full);
			
			// Use it to attach the editor directly to a defined textarea
			WYSIWYG.attach('mesaj'); // default setup
			
			// Use it to display an iframes instead of a textareas
			//WYSIWYG.display('all', full);  
		</script>
<link rel="stylesheet" href="adminstil.css">
<!--#include file="db.asp"-->
</head>
<%
if (Request.QueryString("Blog"))="kaydet" then

id=request.querystring("id")
if isnumeric(id)=false then
response.redirec "giris.asp"
end if

Konu=request.form("Konu")
Mesaj=request.form("Mesaj")
kat_id=request.form("kat_id")
gorunsun=request.form("gorunsun")
yorumdurum=request.form("yorumdurum")


If konu="" or mesaj="" or kat_id="" or yorumdurum="" Then
Response.Write "<center><br><br><center>Bo&#351; B&#305;rakt&#305;&#287;&#305;n&#305;z Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history(back)""><<<GER&#304;</a></center></center>"
Response.End
End if

set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from blog where id="&id&""
blgekle.Open SQL,data,1,3
if not blgekle.eof then

blgekle("mesaj")=mesaj
blgekle("konu")=konu
blgekle("kat_id")=kat_id
blgekle("gorunsun")=gorunsun
blgekle("yorumdurum")=Int(yorumdurum)
blgekle.update
end if

etiket=Trim(Request.Form("etiketler"))
if not etiket="" then
etiketler=Split(etiket,",")

for i=0 to Ubound(etiketler)
Set etk = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from etiket"
etk.Open SQL,data,1,3

etk.Addnew
etk("etiket")=Trim(etiketler(i))
etk("blog_id")=blgekle("id")
etk.update
next
end if

blgekle.Close : Set blgekle = Nothing


Response.Redirect("bloglar.asp")
End if

if (Request.QueryString("Blog"))="duzenle" then

id=request.querystring("id")
if isnumeric(id)=false then
response.redirect "giris.asp"
end if
set blgekle = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from blog where id="&id&""
blgekle.open SQL,data,1,3
if not blgekle.eof then

mesaj=blgekle("mesaj")
if not mesaj="" then
mesaj=Replace(mesaj,"&lt;","&amp;lt;")
mesaj=Replace(mesaj,"&gt;","&amp;gt;")
end if
%>
<body background="images/arka.gif">

<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Blog Düzenle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table class="tablo" width="99%" align="center">
	<form method="post" action="?blog=kaydet&id=<%=id%>">
	<tr>
		<td width="120" align="right"><font class="yazi">Baþlýk :</font></td>
		<td width="1115">
		<input type="text" class="alan" name="konu" size="92" value="<%=blgekle("konu")%>"></td>
	</tr>
	<tr>
		<td width="120" valign="top" align="right"><font class="yazi">Ýçerik :</font></td>
		<td width="1115"><textarea name="mesaj" id="mesaj"><%=mesaj%></textarea></td>
	</tr>
	<tr>
		<td width="120" align="right"><font class="yazi">Etiketler :</font></td>
		<td width="1115">
		<input name="etiketler" type="text" size="92" class="alan"> <font class="yazi" style="font-size:10px">Etiketleri virgül (,) ile ayýrýn. 
		Önceki Etiketlere Ek Yapýlacaktýr.</font></td>
	</tr>
<% set etk=data.execute("Select * from etiket where blog_id="&id&"")
if not etk.eof then %>
	<tr>
		<td></td>
		<td><a href="etiket.asp?id=<%=id%>">Ekli Etiketler:</a> <% Do While Not etk.eof %><a href="etiket.asp?Etiketi=Duzenle&id=<%=etk("id")%>" style="font-weight:normal"><%=etk("etiket")%></a>, <%etk.movenext:loop%></font></td>
	</tr>
<% end if %>
	<tr>
		<td width="120" align="right"><font class="yazi">Kategori :</font></td>
		<td width="1115">
<select name="kat_id" class="alan">
<%
set kate = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from kategori"
kate.open SQL,data,1,3
do while not kate.eof
%>
							<option value="<% = kate("id") %>"<%if blgekle("kat_id")=kate("id") then%> selected<%end if%>><% = kate("ad") %></option>
<%
kate.movenext : loop
if blgekle("gorunsun")=1 then
gorunum="selected"
end if
if blgekle("gorunsun")=2 then
gorunum2="selected"
end if
%>
</select>
		</td>
	</tr>
	<tr>
		<td width="120" align="right"><font class="yazi">Anasayfada :</font></td>
		<td width="1115">
<select size="1" name="gorunsun" class="alan">
<option <%=gorunum%> value="1">Görünsün</option>
<option <%=gorunum2%> value="2">Görünmesin</option>
</select>
		</td>
	</tr>
	<tr>
		<td align="right"><font class="yazi">Yoruma:</font></td>
		<td><input type="radio" name="yorumdurum" value="0"<%if blgekle("yorumdurum")="0" then%> checked<%end if%>><font class="yazi"> Açýk</font>
		<input type="radio" name="yorumdurum" value="1"<%if blgekle("yorumdurum")="1" then%> checked<%end if%>><font class="yazi"> Kapalý</font></td>
	</tr>
	<tr>
		<td width="120">&nbsp;</td>
		<td width="1115">
		<input type="submit" value="Kaydet" class="dugme"></td>
	</tr>
	</form>
</table>
<%
blgekle.Close
Set blgekle = Nothing
Response.End
End if
end if
%>
</body>

</html>
<% end if %>