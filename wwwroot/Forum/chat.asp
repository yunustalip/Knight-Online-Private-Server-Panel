<%@CODEPAGE="1254" LANGUAGE="VbScript" LCID="1055"%>

<%
'*******************************************************
' Kodlar�m� kulland���n�z i�in te�ekk�rler
' Kulland���n�z siteyi bana bildirirseniz sevinirim
' Efkan 
' email :info@aywebhizmetleri.com
' web sayfalar�m� ziyaret etmeyi unutmay�n�z  
' http://www.makineteknik.com
' http://www.binbirkonu.com
' http://www.aywebhizmetleri.com
' http://www.tekrehberim.com
' http://www.hitlinkler.com
' Size uygun bir web sitem mutlaka vard�r ...
' L�TFEN BU T�R �ALI�MALARIN �N�N� KESMEMEK ���N TEL�F YAZILARINI S�LMEY�N
' EME�E SAYGI L�TFEN 
' K���SEL KULLANIM ���N �CRETS�ZD�R D��ER KULLANIMLARDA HAK TALEP ED�LEB�L�R
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
<!-- ODALAR SE� -->
<form name="oda">
<select name="oda" onChange="location=document.oda.oda.options[document.oda.oda.selectedIndex].value;" value="Git">
<option selected value="">Di�er Odalar</option>
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
<B>Chat Odas� : <%=odaadi%></B>


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
Ahlak d��� ifadeler kullanmay�n�z...<BR>
Nick &nbsp;<input type="text" name="uye" size="15" value="<%=Session ("uye")%>" >
<BR>
Mesaj�n�z&nbsp;<input name="mesaj" size="100" >

<BR>
<input type="submit" value="G�nder">
<INPUT TYPE="reset" value="Temizle">

</form>

<A HREF="mailto:info@aywebhizmetleri.com">efkan chat v.1</A>
</td>
</tr>

</table>



