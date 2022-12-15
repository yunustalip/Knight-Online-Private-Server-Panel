<div align="center">
<!--#INCLUDE file="forumayar.asp"-->
<% 
Response.Buffer = True 
%>

<table width="100%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="0">
<tr height="22">
<td background="forumimg/mn.gif" colspan="2" align="center" width="top" width="100%">
<FONT COLOR="#FFFFFF"><B>Son Konu ve Mesajlar </B></FONT></td></tr>

<tr><td align="center" valign="top" width="50%">

<!--  -->


<table width="100%" bgcolor="" bordercolor="#CCFFFF" border="0" cellspacing="0" cellpadding="0">
<%sor = "Select * from sorular where onay=1 order  by id desc"  
forum1.Open sor,forumbag,1,3
If forum1.eof Then 
Else%>
<tr height="20" bgcolor="<%=bgcolor1%>" >
<td width="75%" align="left" valign="center"><B>Son 10 Konu</B></td>
<td width="25%"align="center" valign="center"><B>Ekleyen</B></td>
</tr>

<%
End If
for i=1 to 10
if forum1.eof then exit for
%>
<tr bgcolor="<%=bgcolor2%>" height="20">
<td class="tdbrd" align="left" valign="top">
&nbsp;<A HREF="?part=oku&id=<%=forum1("grp")%>&pid=<%=forum1("altgrp")%>&urun=<%=forum1("id")%>">
<%=kucukharf(forum1("baslik"))%></a></td>

<td class="tdbrd" bgcolor="<%=bgcolor1%>" align="left" valign="top">
<%
'SON MESAJI VEREN ONLÝNE OLUP OLMADIÐI
sor="SELECT * FROM uyeler WHERE id ="&forum1("uyeid")&"  "
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else
zaman=datediff("n",efkan1("sontarih"),Now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>" 
End If
End If
efkan1.close
%>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=forum1("uyeid")%>">
<%=forum1("kadi")%></A><!-- <BR><%=forum1("tarih")%> -->
</td></tr>

<% 
forum1.movenext  'KONULAR DÝZME
Next
forum1.close
%>
</table>


</td><td align="center" valign="top" width="50%">




<!-- SON 5 CEVAP -->


<table   width="100%" bgcolor="" bordercolor="#CCFFFF" border="0" cellspacing="0" cellpadding="0">
<%sor = "Select * from cevaplar where onay=1 order  by id desc"  
forum1.Open sor,forumbag,1,3
If forum1.eof Then 
Else%>
<tr height="20" bgcolor="<%=bgcolor1%>" >
<td width="75%" align="left" valign="center"><B>Son 10 Mesaj</B></td>
<td width="25%"align="center" valign="center"><B>Ekleyen</B></td>
</tr>

<%
End If
for i=1 to 10
if forum1.eof then exit for
%>
<tr bgcolor="<%=bgcolor2%>" height="20">
<td class="tdbrd" align="left" valign="center">
&nbsp;<A HREF="?part=oku&id=<%=forum1("grp")%>&pid=<%=forum1("altgrp")%>&urun=<%=forum1("soruid")%>">
<%=kucukharf(forum1("baslik"))%></a></td>

<td class="tdbrd" bgcolor="<%=bgcolor1%>" align="left" valign="center">
<%
'SON MESAJI VEREN ONLÝNE OLUP OLMADIÐI
sor="SELECT * FROM uyeler WHERE id ="&forum1("uyeid")&"  "
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else
zaman=datediff("n",efkan1("sontarih"),Now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>" 
End If
End If
efkan1.close
%>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=forum1("uyeid")%>">
<%=forum1("kadi")%></A><!-- <BR><%=forum1("tarih")%> -->
</td></tr>

<% 
forum1.movenext  'KONULAR DÝZME
Next
forum1.close
%>
</table>
<!--  -->

</td></tr></table>



<%

set forum =Nothing
set forum1 =nothing
set forum2 =nothing
%>





