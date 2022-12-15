<%@CODEPAGE="1254" LANGUAGE="VbScript" LCID="1055"%>
<!--#INCLUDE file="forumayar.asp"-->
<%
Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage %>


<LINK href="1.css" type=text/css rel=stylesheet>
<HEAD><META HTTP-EQUIV="REFRESH" CONTENT="5; "></HEAD>

<table width="100%" background="" bgcolor="" bordercolor="#CCFFFF" border="0" cellspacing="0" cellpadding="0">
<tr height="350">
<td bgcolor="" align="left" valign="top" width="80%" >
<% 
Response.Buffer = True 

Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath(""&chatyolu&"")
Set efkan = Server.CreateObject("ADODB.Recordset")
Set efkan1 = Server.CreateObject("ADODB.Recordset")


mesajgoster = 8 'PENCEREDEKÝ MESAJ SAYISI

'ONLÝNE CHAT
Session.LCID = 1055
DefaultLCID = Session.LCID 

zamanmiktari=1 
ip=Request.ServerVariables("REMOTE_ADDR")
sor = "SELECT * FROM aktifler where ip='"& ip &"' "
efkan.open sor, sur, 1, 3

If efkan.eof Then
efkan.addnew
efkan("ip")=ip

   If Session("uye")="" Then
efkan("uye")="Misafir"
   else
efkan("uye")=Session ("uye")
   End If
efkan("oda")=Session ("oda")
efkan("tarih")=Now()
Else

   If Session("uye")="" Then
efkan("uye")="Misafir"
   else
efkan("uye")=Session ("uye")
   End If
efkan("oda")=Session ("oda")
efkan("tarih")=Now()
End If
efkan.update
efkan.Close
sor = "SELECT * FROM aktifler"
efkan.open sor, sur, 1, 3
Do While Not efkan.eof 
zaman=datediff("n",efkan("tarih"),now)
if zaman > zamanmiktari then
sor = "DELETE FROM aktifler WHERE  ip = '"&efkan("ip")&"'"
efkan1.open sor, sur, 1, 3
End If
efkan.movenext
Loop
onlineadet = efkan.RecordCount 'ONLÝNE TOPLAM 
efkan.Close 




'10 DAKÝKA ESKÝLERÝ SÝL
sor="SELECT * FROM chat  "
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("n",efkan("tarih"),now) 
if zaman > 10 then
sor="DELETE from chat WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
efkan.Close


'FAZLALARI SÝL BELÝRTÝLENDEN FAZLA MESAJ VARSA 
sor="SELECT  * FROM CHAT where oda ="&session("oda")&" order by id  asc "
efkan.Open sor,Sur,1,3
mesajsayi = efkan.RecordCount
If mesajsayi > mesajgoster then 
i = 0
Do While i =< (mesajgoster -2 ) And Not efkan.Eof
sor="DELETE from chat WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
i = i + 1
efkan.movenext
Loop
End If
efkan.Close


'KAYITLARI DÖKÜYORUM
sor="SELECT * FROM CHAT where oda ="&session("oda")&" order by id asc "
efkan.Open sor,Sur,1,3
If efkan.eof Then
Else
do while not efkan.eof
%>

&nbsp;<FONT  COLOR="#3300FF"><B><%=kucukharf(efkan("uye"))%>:</B></FONT>
<I><FONT SIZE="1" COLOR="#CC99CC"><%=Right(efkan("tarih"),8)%></FONT></I>
<%=kucukharf(efkan("mesaj"))%><P>

<% efkan.movenext 
loop
End If
efkan.close
%>

</td>




<td  bgcolor="" align="left" valign="top" width="20%" >

<fieldset>
<legend>
<B>Bu odada aktifler</B><BR>
</legend>
<% 'ODADAKÝLERÝ DÖK 
sor="SELECT  * FROM aktifler where oda ="&session("oda")&"  order by id desc "
efkan.Open sor,Sur,1,3
If efkan.eof Then
Else
do while not efkan.eof %>

<FONT COLOR="blue"><B><%=kucukharf(efkan("uye"))%></B></FONT><BR>
<% efkan.movenext 
loop
efkan.close
End If
%>
</fieldset>

<P>
<fieldset><legend>
<B>Chat Odalarýmýz</B><BR>
</legend>
<% 'ODALARDAKÝ ZÝYARETCÝ SAYILARI
sor = "SELECT * FROM oda "
efkan.open sor, sur, 1, 3
Do While Not efkan.eof 
sor="SELECT  * FROM aktifler where oda ="&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
adet=efkan1.RecordCount %>
<%=efkan("adi")%> :<B><%=adet%> </B><BR>
<%efkan1.close
efkan.movenext
Loop 
efkan.close%>
</fieldset>

</td></tr></table>

<%
Set efkan1=Nothing
Set efkan=Nothing
%>