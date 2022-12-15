<%@CODEPAGE="1254" LANGUAGE="VbScript" LCID="1055"%>

<%
'*******************************************************
' Kodlarýmý kullandýðýnýz için teþekkürler
' Kullandýðýnýz siteyi bana bildirirseniz sevinirim
' Efkan 
' email :info@aywebhizmetleri.com
' web sayfalarýmý ziyaret etmeyi unutmayýnýz  
' http://www.makineteknik.com
' http://www.binbirkonu.com
' http://www.aywebhizmetleri.com
' http://www.tekrehberim.com
' http://www.hitlinkler.com
' Size uygun bir web sitem mutlaka vardýr ...
' LÜTFEN BU TÜR ÇALIÞMALARIN ÖNÜNÜ KESMEMEK ÝÇÝN TELÝF YAZILARINI SÝLMEYÝN
' EMEÐE SAYGI LÜTFEN 
' KÝÞÝSEL KULLANIM ÝÇÝN ÜCRETSÝZDÝR DÝÐER KULLANIMLARDA HAK TALEP EDÝLEBÝLÝR
'*******************************************************
%>




<%
session.TimeOut = 60 
Server.ScriptTimeOut = 60

Response.Buffer = True

tema=request.querystring("tema")
If tema<>"" then
session("tema")=tema
ElseIf session("tema")="" then
session("tema")=1
ElseIf session("tema")="" then
session("tema")=session("tema")
End If


%>


<!--#INCLUDE file="forumayar.asp"-->

<body bgcolor="<%=bgcolor2%>">
<div align="center">
<HTML>
<HEAD>
<TITLE><%=title%></TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9">
<meta http-equiv="content-language" content="tr">
<meta name="author" content="Efkan Ay">
<META NAME="description" CONTENT="Forum Sayfanýz">
<meta name="keywords" content="<%=KEYWORDS%>"> 
<meta http-equiv="revisit-after" content="2 days">
</HEAD>



<table  class="tborder"  width="800" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr height="80">
<td width="100%" align="center" valign="center">
<IMG SRC="logom.gif" WIDTH="800" HEIGHT="80" BORDER="0" ALT="">
</td></tr>



<tr><td bgcolor="<%=bgcolor1%>" width="100%" align="center" valign="center">
<!--#INCLUDE file="uyemenu.asp"-->



<!-- ARAMA VE HIZLI MENU -->
<table  width="100%" bgcolor="<%=bgcolor2%>" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr height="30"><td width="40%" align="left" valign="center">
&nbsp;<A HREF="default.asp"><B>Ana Sayfa</B></A>

|


<a href="chat.asp" onClick="window.name='ana'; window.open('chat.asp','new', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no, resizable=no,copyhistory=no,width=800,height=600'); return false;" ><B>Chat</B>
<%Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath(""&chatyolu&"")
Set efkan = Server.CreateObject("ADODB.Recordset")
Set efkan1 = Server.CreateObject("ADODB.Recordset")
sor = "SELECT * FROM aktifler"
efkan.open sor, sur, 1, 3
Do While Not efkan.eof 
zaman=datediff("n",efkan("tarih"),now)
if zaman > 1 then
sor = "DELETE FROM aktifler WHERE  ip = '"&efkan("ip")&"'"
efkan1.open sor, sur, 1, 3
End If
efkan.movenext
Loop
adet = efkan.RecordCount 'ONLÝNE TOPLAM 
efkan.Close 
If adet=0 Then 
Else %>
<IMG SRC="images/yanson1.gif" WIDTH="12" HEIGHT="12" BORDER="0" ALT=""><B><%=adet%>kiþi</B>
<%End If%></a>

|

<A HREF="?part=map1"><B>Site Haritasý</B></A>
|
<A HREF="?part=kurallar"><B>Kurallar</B></A>
</td>
<FORM method="POST" action="default.asp?part=ara">
<td width="20%" align="center" valign="center">
<input type="text" name="ara" size="15" maxlength="15">&nbsp;
<select name="nerde">
<option value="soru">Konularda</option>
<option value="cevap">Mesajlarda</option>
</select>&nbsp;<input type="submit" value=" Ara ">
</td></FORM>
<form name="jump">
<td width="20%" align="center" valign="center"><!--#INCLUDE file="map.asp"--></td></form>

<form name="tema">
<td width="10%" align="center" valign="center">
<select name="tema" onChange="location=document.tema.tema.options[document.tema.tema.selectedIndex].value;" value="Git">
<option selected value="">Tema</option>
<option value="default.asp?tema=1">1</option>
<option value="default.asp?tema=2">2</option>
<option value="default.asp?tema=3">3</option>
<option value="default.asp?tema=4">4</option>
<option value="default.asp?tema=5">5</option>
</select>
</td></FORM></tr>
</table>
<!--  -->




</td></tr>





<tr height="500">
<td bgcolor="<%=bgcolor2%>" width="100%" align="center" valign="top">
<!--#INCLUDE file="part.asp"-->
</td></tr>

<tr>
<td width="100%" align="center" valign="top">
<!--#INCLUDE file="ist.asp"-->
<!--#INCLUDE file="online.asp"-->
</td></tr>


<%

'///////// EÐER BÝRAZ EMEÐE SAYGINIZ VARSA SÝLMESSÝNÝZ //////////////////
'///////// BU SCRÝPT ÝÇÝN ÇOK EMEK HARCANMIÞTIR VE HARCANMAYA DEVAM EDECEKTÝR  //////////////////
'///////// EÐER BÝRAZ EMEÐE SAYGINIZ VARSA SÝLMESSÝNÝZ //////////////////
sHTML = sHTML & "<tr bgcolor="&bgcolor2&" height=""30"">"
sHTML = sHTML & "<td width=""100%"" align=""center"" valign=""center""><FONT SIZE=""1"" >"
sHTML = sHTML & "<A HREF=""http://www.aywebhizmetleri.com"" target=""_blank"">efkan forum v.4.3</A>&nbsp;"
sHTML = sHTML & "<A HREF=""mailto:info@aywebhizmetleri.com"">Tasarým Kodlama Efkan Ay</A>&nbsp;&copy;2006"
sHTML = sHTML & "</FONT></td></tr>"
Response.Write shtml
'///////// EÐER BÝRAZ EMEÐE SAYGINIZ VARSA SÝLMESSÝNÝZ //////////////////
'///////// BU SCRÝPT ÝÇÝN ÇOK EMEK HARCANMIÞTIR VE HARCANMAYA DEVAM EDECEKTÝR  //////////////////
'///////// EÐER BÝRAZ EMEÐE SAYGINIZ VARSA SÝLMESSÝNÝZ //////////////////
%>

</table>

