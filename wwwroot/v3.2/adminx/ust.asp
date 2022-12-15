<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Blog</title>
<style>
body
a:active
{
font-size: 11px;
font-family: tahoma;
font-weight:bold;
color: #000000;
text-decoration: none;
}
a:link
{
font-size: 11px;
font-family: tahoma;
font-weight:bold;
color: #000000;
text-decoration: none;
}
a:visited
{
font-size: 11px;
font-family: tahoma;
font-weight:bold;
color: #000000;
text-decoration: none;
}
a:hover
{
font-size: 11px;
font-family: tahoma;
font-weight:bold;
color: #3F5D38;
text-decoration: none;
}


.blog a:active
{
font-size: 11px;
font-family: tahoma;
color: #000000;
font-weight:bold;
height:21px; width:100%; padding-top:3px; padding-bottom:1px;
}
.blog a:link
{
font-size: 11px;
font-family: tahoma;
color: #000000;
font-weight:bold;
height:21px; width:100%; padding-top:3px; padding-bottom:1px;
}
.blog a:visited
{
font-size: 11px;
font-family: tahoma;
color: #000000;
font-weight:bold;
height:21px; width:100%; padding-top:3px; padding-bottom:1px;
}
.blog a:hover
{
font-size: 11px;
font-family: tahoma;
font-weight:bold;
color: #000000;
text-decoration: none;
height:21px; width:100%; padding-top:3px; padding-bottom:2px;
background-color: #FFD38E; background-image: url('images/ic_bg.gif'); background-repeat: repeat-x;border:1px solid #3F5D38;
}
</style>
</head>

<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" background="images/arka.gif">
<center>
<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0" height="25" background="images/bg.gif">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="88" align="center"><font class="blog"><a href="#" onMouseover="showit(0)">&nbsp;&nbsp;&nbsp;&nbsp;Bloglar&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="#" onMouseover="showit(1)">&nbsp;&nbsp;&nbsp;&nbsp;Galeri&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="#" onMouseover="showit(2)">&nbsp;&nbsp;&nbsp;&nbsp;Yorumlar&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="#" onMouseover="showit(3)">&nbsp;&nbsp;&nbsp;&nbsp;Mesajlar&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="#" onMouseover="showit(4)">&nbsp;&nbsp;&nbsp;&nbsp;İletişim&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="anket.asp" onMouseover="showit(5)" target="alt">&nbsp;&nbsp;&nbsp;&nbsp;Anketler&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="ayarlar.asp" onMouseover="showit(5)" target="alt">&nbsp;&nbsp;&nbsp;&nbsp;Ayarlar&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="_hakkimda.asp" onMouseover="showit(5)" target="alt">&nbsp;&nbsp;&nbsp;&nbsp;Hakkımda&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td width="4" background="images/ayrac.gif" align="center"></td>
		<td width="88" align="center"><font class="blog"><a href="alt.asp" onMouseover="showit(5)" target="alt">&nbsp;&nbsp;&nbsp;&nbsp;İstatistik&nbsp;&nbsp;&nbsp;&nbsp;</a></font></td>
		<td align="right"><a href="ust.asp?admin=cikis" onMouseover="showit(5)" target="_top">ÇIKIŞ YAP</a></td>
		<td width="16"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
</center>
<div align="center">
<table border="0" width="99%" height="21" id="table2" cellpadding="0" style="border:1px solid #3F5D38; border-collapse: collapse; margin-top:1px; margin-bottom:1px" bgcolor="#F8F8F8">
	<tr>
		<td width="95%" style="padding-left: 5px" height="23"><div id="describe" onMouseover="clear_delayhide()"><b><font face="Tahoma" style="font-size: 11px">Menü Seçiniz</font></b></div></td>
	</tr>
</table>
</div>
</body>

<script language="JavaScript1.2">
var submenu=new Array()
submenu[0]='<a href="ekle.asp" target="alt">Blog Ekle</a> • <a href="bloglar.asp" target="alt">Bloglar</a> • <a href="kategori_ekle.asp" target="alt">Kategori Ekle</a> • <a href="kategoriler.asp" target="alt">Kategoriler</a> • <a href="etiket.asp" target="alt">Etiketler</a> • <a href="linkler.asp" target="alt">Bağlantılar</a> • <a href="linkler.asp?link=ekle" target="alt">Bağlantı Ekle</a>'
submenu[1]='<a href="resimekle.asp" target="alt">Resim Ekle</a> • <a href="resimler.asp" target="alt">Resimler</a> • <a href="g_kategori_ekle.asp" target="alt">Albüm Ekle</a> • <a href="g_kategoriler.asp" target="alt">Albümler</a>'
submenu[2]='<a href="_yorum.asp" target="alt">Yorum Yönetimi</a> • <a href="yorum_onay.asp" target="alt">Onay Bekleyen Yorumlar</a>'
submenu[3]='<a href="z_d.asp" target="alt">Ziyaretçi Defteri Yönetimi</a> • <a href="zd_onay.asp" target="alt">Onay Bekleyen Mesajlar</a>'
submenu[4]='<a href="ileti.asp?ileti=yanitla&id=0" target="alt">Mail Yaz/Gönder</a> • <a href="ileti.asp" target="alt">İletiler</a>'
submenu[5]=''

var delay_hide=500
var menuobj=document.getElementById? document.getElementById("describe") : document.all? document.all.describe : document.layers? document.dep1.document.dep2 : ""

function showit(which){
clear_delayhide()
thecontent=(which==-1)? "" : submenu[which]
if (document.getElementById||document.all)
menuobj.innerHTML=thecontent
else if (document.layers){
menuobj.document.write(thecontent)
menuobj.document.close()
}
}

function resetit(e){
if (document.all&&!menuobj.contains(e.toElement))
delayhide=setTimeout("showit(-1)",delay_hide)
else if (document.getElementById&&e.currentTarget!= e.relatedTarget&& !contains_ns6(e.currentTarget, e.relatedTarget))
delayhide=setTimeout("showit(-1)",delay_hide)
}

function clear_delayhide(){
if (window.delayhide)
clearTimeout(delayhide)
}

function contains_ns6(a, b) {
while (b.parentNode)
if ((b = b.parentNode) == a)
return true;
return false;
}
</script>
</html>
<% End if %>
<%
if (Request.QueryString("admin"))="cikis" then
session("admin")=""
session.abandon
response.redirect("admin.asp")
End if
%>