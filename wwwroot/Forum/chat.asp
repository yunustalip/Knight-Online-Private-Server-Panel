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

<!--#INCLUDE file="forumayar.asp"-->


<% Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage %>


<body  onLoad="self.focus();document.formcevap.mesaj.focus()">



<LINK href="1.css" type=text/css rel=stylesheet>
<div align="center">
<table width="100%" background="" bgcolor="" bordercolor="#CCFFFF" border="0" cellspacing="0" cellpadding="0">

<tr height="">
<td align="center" valign="center" width="100%">

<% 

Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath(""&chatyolu&"")
Set efkan = Server.CreateObject("ADODB.Recordset")
Set efkan1 = Server.CreateObject("ADODB.Recordset")

oda=request.querystring("oda")
If oda<>"" then
session("oda")=oda
ElseIf session("oda")="" then
session("oda")=1
ElseIf session("oda")="" then
session("oda")=session("oda")
End If %>


<BR><BR>
<!-- ODALAR SEÇ -->
<form name="oda">
<select name="oda" onChange="location=document.oda.oda.options[document.oda.oda.selectedIndex].value;" value="Git">
<option selected value="">Diðer Odalar</option>
<% sor = "SELECT * FROM oda "
efkan.open sor, sur, 1, 3
Do While Not efkan.eof %>
<option value="?part=chat&oda=<%=efkan("id")%>"><%=efkan("adi")%></option>
<% efkan.movenext
Loop 
efkan.close%>
</select></FORM>



<% sor = "SELECT * FROM oda where id ="&session("oda")&"  "
efkan.open sor, sur, 1, 3 
odaadi=efkan("adi")
efkan.close%>
<B>Chat Odasý : <%=odaadi%></B>


</td></tr>


<tr height="350">
<td align="center" valign="center" width="100%" >

<% if Request.form("uye")<>"" And  Request.form("mesaj")<>""   then
sor="SELECT * FROM CHAT   "
efkan.Open sor,Sur,1,3
efkan.AddNew

efkan("uye")       =temizle(Request.Form ("uye"))
mesaj                =left(Request.Form ("mesaj"),500)
efkan("mesaj")    =temizle(mesaj)
efkan("tarih")      =Now()
efkan("oda")       = session("oda")
efkan("ip")          = Request.ServerVariables("REMOTE_ADDR")
efkan.update
Session ("uye")    = efkan("uye")
efkan.close
End If 
%>
<bgsound src="ses.wma" loop="1">



<iframe  name="orta" width="100%" height="350" frameborder="0" scrolling="no" src="chat1.asp" noresize>
</iframe>


</td>



</tr><tr height="70"><td align="center" valign="center" width="100%" >

<form name="formcevap"  action="chat.asp" method="post" >
<BR>
Ahlak dýþý ifadeler kullanmayýnýz...<BR>
Nick &nbsp;<input type="text" name="uye" size="15" value="<%=Session ("uye")%>" >
<BR>
Mesajýnýz&nbsp;<input name="mesaj" size="100" >

<BR>
<input type="submit" value="Gönder">
<INPUT TYPE="reset" value="Temizle">

</form>

<A HREF="mailto:info@aywebhizmetleri.com">efkan chat v.1</A>
</td>
</tr>

</table>



