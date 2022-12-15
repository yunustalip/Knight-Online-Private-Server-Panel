<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Blog Ekle</title>
<link rel="stylesheet" href="adminstil.css">
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
</head>
<!--#include file="db.asp"-->
<%
Function temizle(trh)
	if trh<10 then
		trh=Right(trh,1)
	end if
temizle=trh
End Function

Function Duzelt(duzgun)
if mid(duzgun,3,1)="." then
x=split(duzgun,".")
	ay=x(1)
	gun=x(0)
	yil=x(2)
duzgun= temizle(ay) & "/" & temizle(gun) & "/" & yil
end if
duzelt=duzgun
End Function

if (Request.QueryString("Blog"))="ekle" then

Konu=request.form("Konu")
Mesaj=request.form("Mesaj")
kat_id=request.form("kat_id")
gorunsun=request.form("gorunsun")
yorumdurum=request.form("yorumdurum")

If konu="" or mesaj="" or kat_id="" or yorumdurum="" Then
Response.Write "<center>Boş Bıraktığınız Alan Var.<br> Geri Dönüp Doldurunuz.<br><a href=""javascript:history(back)""><<<GERİ</a></center>"
Response.End
End if

Set blgekle = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from blog"
blgekle.Open SQL,data,1,3

blgekle.Addnew
blgekle("Mesaj")=Mesaj
blgekle("Konu")=Konu
blgekle("kat_id")=kat_id
blgekle("Tarih")=duzelt(date)
blgekle("gorunsun")=gorunsun
blgekle("yorumdurum")=Int(yorumdurum)
blgekle.update

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

blgekle.close : set blgekle=nothing

Response.Redirect("?Blog=Eklendi")
End if

if (Request.QueryString("Blog"))="Eklendi" then
response.write("<script>alert('Blog Başarıyla Eklendi');</script>")
End if
%>
<body background="images/arka.gif">

<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Blog Ekle</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table class="tablo" width="99%" align="center">
<form action="?Blog=ekle" method="post">
	<tr>
		<td width="120" align="right"><font class="yazi">Başlık :</font></td>
		<td width="1115">
		<input type="text" class="alan" name="konu" size="92"></td>
	</tr>
	<tr>
		<td width="120" valign="top" align="right"><font class="yazi">İçerik :</font></td>
		<td width="1115"><textarea name="mesaj" id="mesaj"></textarea></td>
	</tr>
	<tr>
		<td width="120" align="right"><font class="yazi">Etiketler :</font></td>
		<td width="1115">
		<input name="etiketler" type="text" size="92" class="alan"> <font class="yazi" style="font-size:10px">Etiketleri virgül (,) ile ayırın. 
		</font></td>
	</tr>
	<tr>
		<td width="120" align="right"><font class="yazi">Kategori :</font></td>
		<td width="1115">
<select name="kat_id" class="alan">
<option selected value="">Kategori seçin</option>
<%
set katsec = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from kategori"
katsec.open SQL,data,1,3
Do While Not katsec.EOF
%>
<option value="<%=katsec("id")%>"><%=katsec("ad")%></option>
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
		<td width="120" align="right"><font class="yazi">Anasayfada :</font></td>
		<td width="1115">
<select size="1" name="gorunsun" class="alan">
<option selected value="1">Görünsün</option>
<option value="2">Görünmesin</option>
</select>
		</td>
	</tr>
	<tr>
		<td align="right"><font class="yazi">Yoruma:</font></td>
		<td><input type="radio" name="yorumdurum" value="0" checked><font class="yazi"> Açık</font>
		<input type="radio" name="yorumdurum" value="1"><font class="yazi"> Kapalı</font></td>
	</tr>
	<tr>
		<td width="120">
		</td>
		<td width="1115">
		<input type="submit" value="Kaydet" class="dugme"></td>
	</tr>
</form>
</table>
</body>

</html>
<% End if %>