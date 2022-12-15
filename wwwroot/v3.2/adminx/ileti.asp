<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Admin Paneli</title>
<link rel="stylesheet" href="adminstil.css">
<% if (Request.QueryString("ileti"))="yanitla" then %>
		<script type="text/javascript" src="scripts/wysiwyg.js"></script>
		<script type="text/javascript" src="scripts/wysiwyg-settings.js"></script>
		<!-- 
			Attach the editor on the textareas
		-->
		<script type="text/javascript">
			// Use it to attach the editor to all textareas with full featured setup
			//WYSIWYG.attach('all', full);
			
			// Use it to attach the editor directly to a defined textarea
			WYSIWYG.attach('ileti'); // default setup
			
			// Use it to display an iframes instead of a textareas
			//WYSIWYG.display('all', full);  
		</script>
<% end if %>
</head>
<!--#include file="../ayar.asp"-->
<!--#include file="db.asp"-->
<!--#include file="../filtre.asp"-->
<body background="images/arka.gif">
<%
if (Request.QueryString("ileti"))="sil" then
id=request.querystring("id")
data.Execute("DELETE FROM iletisim where id like '"&id&"'")
Response.Redirect "ileti.asp"

elseif (Request.QueryString("ileti"))="yanitla" then
id=Request.QueryString("id")
if not isnumeric(id)=false or id="" then
set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from iletisim where id="&id&""
rs.open SQL,data,1,3
if not rs.eof then
	mail=rs("mail")
	konu="YNT: "&rs("konu")
	yer=rs("yer")
	ileti="<br><br><br><br><blockquote style=""PADDING-LEFT: 5px; MARGIN-LEFT: 15px; BORDER-LEFT: #000000 2px solid""><b>Alıntı: '"&rs("isim")&"'<br>"&rs("tarih")&"</b><br>"&rs("mesaj")&"</blockquote>"
	web=rs("url")
end if
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function validate(form) {
if (form.gonderen.value == "") {
   alert("E-mail Adresinizi Yazın");
   return false; }
if (form.alici.value == "") {
   alert("Alıcı Mail Adresini Yazın");
   return false; }
if (form.konu.value == "") {
   alert("Başlık Girin");
   return false; }
return true;
}
</SCRIPT>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Mail Yaz</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
<table width="99%" border="0" id="table1" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="?ileti=gonder&id=<%=id%>" method="post" onSubmit="return validate(this)">
	<tr>
		<td width="295">
		<p align="right"><font class="yazi">Admin&nbsp;Mail: </font></td>
		<td width="188"><input type="text" name="gonderen" size="36" class="alan"></td>
		<td width="146">
		<p align="right"><font class="yazi">Alıcı: </font></td>
		<td width="598"><input type="text" name="alici" value="<%=mail%>" size="41" class="alan"></td>
	</tr>
	<tr>
		<td width="295">
		<p align="right"><font class="yazi">Konu: </font></td>
		<td width="188"><input type="text" name="konu" value="<%=konu%>" size="36" class="alan"></td>
		<td width="146" align="right"><%if not rs.eof then%><font class="yazi">Web / Yer: </font><% end if %></td>
		<td width="598"><%if not rs.eof then%><font class="yazi">&nbsp;&nbsp; <a href="<%=web%>" target="_blank"><%=web%></a> / <%=yer%></font><% end if %></td>
	</tr>
	<tr>
		<td width="295">
		<p align="right"><font class="yazi">İleti: </font></td>
		<td colspan="3">
		<textarea name="ileti" id="ileti"><%=ileti%></textarea>
		</td>
	</tr>
	<% if not rs.eof then %>
	<tr>
		<td></td>
		<td colspan="3"><input type="checkbox" name="sil"><font class="yazi">Mail Gönderildikten Sonra Bu İletiyi Sil</font></td>
	</tr>
	<% end if %>
	<tr>
		<td width="295"></td>
		<td colspan="3"><input type="submit" value="Gönder" class="dugme"></td>
	</tr>
</form>
	<tr>
		<td colspan="4" align="center"><font class="yazi">Not: Admin mail adresiniz domain adınıza bağlı bir adres olmalıdır aksi halde 
		<u>Önemsiz/Junk</u> Mail Olarak Gidebilir.</font></td>
	</tr>
</table>
</div>
</table>
<%
rs.close : set rs=nothing
end if
elseif (Request.QueryString("ileti"))="gonder" then
konu=Request.Form("konu")
gonderen=Request.Form("gonderen")
alici=Request.Form("alici")
ileti=Request.Form("ileti")
id=Request.QueryString("id")

if ileti="" or konu="" or gonderen="" or alici="" then
	response.write "Boş Bıraktığınız Alan Var <a href=""javascript:history.back()"">&lt;&lt;Geri</a> Gidip Boş Bıraktığınız Alanları Doldurun"
else

if StrMail="1" then
Set Mail=Server.CreateObject("CDONTS.NewMail")
Mail.To=alici
Mail.MailFormat = 1
else
Set Mail=Server.CreateObject("Jmail.Message")
Mail.Charset = "ISO-8859-9"
Mail.AddRecipient alici
Mail.contenttype="text/html"
end if
Mail.From=gonderen
Mail.Subject=konu
eb=ileti
if strmail<>1 then
Mail.htmlbody=eb
else
Mail.Body=eb
end if
Mail.Send(strmailserver)
if Request.Form("sil")="on" then
if not isnumeric(id)=false then
data.Execute("DELETE FROM iletisim where id="&id&"")
end if
end if
response.write "<script>alert('Mail Başarıyla Gönderildi')</script>"
response.write "<script>location.href('ileti.asp');</script>"
end if
elseif (Request.QueryString("ileti"))="oku" then
id=Request.QueryString("id")
if not isnumeric(id)=false or id="" then
set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "select * from iletisim where id="&id&""
rs.open SQL,data,1,3
if not rs.eof then
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">İleti Oku</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo" align="center">
	<tr>
		<td class="ust" height="24" style="padding:4px" align="right"><a href="?ileti=yanitla&id=<%=id%>">YANITLA</a> - <a href="?ileti=sil&id=<%=id%>" onclick="return confirm('Silmek İstediğinizden Eminmisiniz?');">SİL</a></td>
	</tr>
	<tr>
		<td class="tablo1" height="24" style="padding:4px"><font class="yazi">İsim: <font style="font-weight:normal"><%=rs("isim")%></font>,&nbsp;&nbsp;&nbsp;&nbsp; Yer: <font style="font-weight:normal"><%=rs("yer")%></font>,&nbsp;&nbsp;&nbsp;&nbsp; Web: <a href="<%=rs("url")%>" target="_blank" style="font-weight:normal"><%=rs("url")%></a>,&nbsp;&nbsp;&nbsp;&nbsp; Tarih: <font style="font-weight:normal"><%=rs("tarih")%></font>,&nbsp;&nbsp;&nbsp;&nbsp; E-mail: <font style="font-weight:normal"><%=rs("mail")%></font></font></td>
	</tr>
	<tr>
		<td style="padding:4px"><font class="baslik"><%=rs("konu")%><br></font><font class="yazi" style="font-weight:normal"><%=rs("mesaj")%></font></td>
	</tr>
</table>
<%
end if
end if
else
siralama=filtre(Request.QueryString("siralama"))
tip=filtre(Request.QueryString("tip"))
sayi=filtre(Request.QueryString("sayi"))

	if siralama="1" then
		sira="konu"
	elseif siralama="2" then
		sira="hit"
	elseif siralama="3" then
		sira="id"
	else
		sira="id"
	end if

	if tip="1" then
		t="desc"
	elseif tip="2" then
		t="asc"
	else
		t="desc"
	end if

	if sayi="1" then
		s="25"
	elseif sayi="2" then
		s="50"
	elseif sayi="3" then
		s="100"
	else
		s="25"
	end if
%>
<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">İletiler</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
<table border="0" width="99%" id="table2" style="border-collapse: collapse" class="tablo">
	<tr>
		<td height="24" width="43" class="ust"><font class="yazi">ID</font></td>
		<td height="24" width="300" class="ust"><font class="yazi">Kimden</font></td>
		<td height="24" width="300" class="ust"><font class="yazi">Konu</font></td>
		<td height="24" width="310" class="ust"><font class="yazi">Tarih</font></td>
		<td height="24" width="130" class="ust"><font class="yazi">İşlem</font></td>
	</tr>
	
<% response.buffer = "true" %>

<%
Set zd_msg = Server.CreateObjecT("ADODB.recordSet")
rSQL = "Select * from iletisim order by "&sira&" "&t&""
zd_msg.open rSQL,data,3,3
adet = zd_msg.recordcount
if not zd_msg.eof then

sayfa = Request.QueryString("sayfa")
    if isnumeric(sayfa)=false then
        Response.redirect "index.asp"
    Else
if sayfa="" then sayfa=1
zd_msg.pagesize = (s)
sayfa_sayisi = zd_msg.pagecount
if Cint(sayfa)<1 then sayfa=1
if Cint(sayfa_sayisi) < Cint(sayfa) then sayfa=sayfa_sayisi
zd_msg.absolutepage = sayfa
mode = 2
for i=1 to zd_msg.pagesize
if zd_msg.eof then
exit for
end if
	if mode=1 then
	stil="tablo1"
	else
	stil=""
	end if
%>
	<tr>
		<td width="43" class="<%=stil%>" height="24"><font class="yazi"><%=zd_msg("id")%></font></td>
		<td width="300" class="<%=stil%>" height="24"><font class="yazi"><a href="?ileti=oku&id=<%=zd_msg("id")%>"><%=zd_msg("isim")%></a></td>
		<td width="390" class="<%=stil%>" height="24"><font class="yazi"><a href="?ileti=oku&id=<%=zd_msg("id")%>"><%=zd_msg("konu")%></a></font></td>
		<td width="220" class="<%=stil%>" height="24"><font class="yazi"><%=FormatDateTime(zd_msg("tarih"),1)%>&nbsp;<%=FormatDateTime(zd_msg("tarih"), 4)%></font></td>
		<td width="130" class="<%=stil%>" height="24"><font class="yazi"><a href="?ileti=yanitla&id=<%=zd_msg("id")%>">Yanıtla</a> - <a href="?ileti=sil&id=<%=zd_msg("id")%>">Sil</a></td>
	</tr>
<%
zd_msg.movenext
	if mode=2 then
	mode=1
	else
	mode=2
	end if
%>
<% next %> 
	</table>
</div>
<div align="center">
	<table border="0" width="99%" id="table2" cellpadding="0" style="border-collapse: collapse" class="tablo">
<form action="" method="get">
		<tr>
			<td>
		<select name="siralama" size="1" class="alan">
        <option value="1"<%if sira="konu" then%> selected<%End if%>>İsme Göre</option>
        <option value="2"<%if sira="hit" then%> selected<%End if%>>Hite Göre</option>
        <option value="3"<%if sira="id" then%> selected<%End if%>>Tarihe Göre</option>
        </select><select name="tip" size="1" class="alan">
        <option value="1"<%if t="desc" then%> selected<%End if%>>Artan</option>
        <option value="2"<%if t="asc" then%> selected<%End if%>>Azalan</option>
        </select><select name="sayi" size="1" class="alan">
        <option value="1"<%if s="25" then%> selected<%End if%>>25</option>
        <option value="2"<%if s="50" then%> selected<%End if%>>50</option>
        <option value="3"<%if s="100" then%> selected<%End if%>>100</option>
        </select><input type="submit" value="Sırala" class="dugme">
			</td>
		</tr>
</form>
	</table>

	<table width="99%" border="0" class="tablo" id="table1" cellpadding="0" style="border-collapse: collapse">
		<tr>
			<td colspan="3" align="center"><font class="yazi">Toplam <%=adet%> kayıt, <%=sayfa_sayisi%> Sayfada Gösterilmektedir.</font></td>
			</tr>
		<tr>
			<td align="center" valign="center">
<%
If sayfa > 1 Then
response.write "<b><a href=""?sayfa=1&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""ilk sayfa"">««</a></b> "
a = sayfa -1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""Önceki"">«</a></b> "
End If
If sayfa + 10 > sayfa_sayisi Then
b = sayfa_sayisi 
Elseif sayfa_sayisi < 10 Then
sayfa_sayisi = 1
Else
b = sayfa + 10
End If
If sayfa < 10 Then
c = 1
Else
c = sayfa - 10
End If
if c < 1 then 
c = 1
end if
For j = c To b
If j = CInt(sayfa) Then
Response.Write "<font class=""yazi"">[<b>" & j & "</b>] </font>"
Else
Response.Write "<b><a href=""?sayfa=" & j & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""">" & j & "</a></b> "
End If
Next
if Cint(sayfa) < sayfa_sayisi then
a = sayfa + 1
Response.Write "<b><a href=""?sayfa=" & a & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""Sonraki"">»</a></b> "
Response.Write "<b><a href=""?sayfa=" & sayfa_sayisi & "&siralama="&siralama&"&tip="&tip&"&sayi="&sayi&""" title=""Son Sayfa"">»»</a></b>"
End If
zd_msg.close : set zd_msg = nothing
%>
			</td>
		</tr>
	</table>
</div>
<%
End if
Else
response.write "<tr>"&chr(10)
response.write "<td colspan=""5"" height=""24""><font class=""yazi""><center>Kayıt Bulunamadı</center></font></td>"&chr(10)
response.write "</tr>"&chr(10)
response.write "</table>"
End if
End if
end if
%>