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



<div align="center">

<!-- KATEGORÝLER DÝZÝLÝYOR -->
<table  width="100%" bgcolor="" bordercolor="#333399" border="0" cellspacing="0" cellpadding="0">
<tr>
<td  background="forumimg/mn.gif" width="3%" height="38" align="center" valign="center"></td>
<td  background="forumimg/mn.gif" width="40%" height="38" align="left" valign="center">
<FONT COLOR="#FFFFFF"><B>Forum Kategorileri</B></FONT></td>
<td  background="forumimg/mn.gif" width="5%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Konu</B></FONT></td>
<td  background="forumimg/mn.gif" width="5%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Mesaj</B></FONT></td>
<td background="forumimg/mn.gif" width="30%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Son Mesaj</B></FONT></td>
</tr>


<!-- ANA KATLAR -->
<%sor = "Select * from grup where altgrp=0 order  by grp asc"  
forum.Open sor,forumbag,1,3
if forum.eof or forum.bof then
Response.Write "<BR><BR><BR><center><B>Kayýt yok</B><P>"
Response.End
End If
do while not forum.eof  
%>


<tr height="25">
<td border="0" background="forumimg/mn.gif" colspan="5" width="100%" align="left" valign="center" >
<A onclick="javascript:toggle('<%=forum("id")%>');return false;" href="#"><IMG SRC="images/yuk.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT=""></A>
&nbsp;
<A HREF="?part=grp&id=<%=forum("id")%>">
<B><FONT COLOR="#FFFFFF"><%=kucukharf(forum("grp"))%></FONT></B></A>

&nbsp;&nbsp;&nbsp;
<%'BU KATTA KAÇ KÝÞÝ VAR
sor = "SELECT * FROM online where grp="& forum("id") &"  "
efkan.open sor, sur, 1, 3
adet=efkan.recordcount
efkan.Close
Response.Write "<FONT COLOR=""#FFFFFF"">"
Response.Write "Aktif<IMG SRC=""images/tekil.gif"" WIDTH=""17""  BORDER=""0"">&nbsp;"
Response.Write adet
Response.Write "</FONT>"
%>

</td></tr>

<tr><td colspan="5" id="<%=forum("id")%>" style="DISPLAY: yes" width="100%">


<!-- ALT KATLAR DÝZÝLÝYOR-->
<table  width="100%" bgcolor="" bordercolor="#333399" border="0" cellspacing="0" cellpadding="3">
<%sor = "Select * from grup where altgrp="&forum("id")&" order  by grp asc"  
forum1.Open sor,forumbag,1,3
do while not forum1.eof  %>
<tr height="25">
<td width="3%" class="tdbrd"  bgcolor="<%=bgcolor1%>" background="" align="center" valign="center">
<IMG SRC="forumimg/haber.gif" WIDTH="24" HEIGHT="24" BORDER="0" ALT="">
</td>

<td width="40%" class="tdbrd"  bgcolor="<%=bgcolor2%>" background="" align="left" valign="center">
<A HREF="?part=altgrp&id=<%=forum("id")%>&pid=<%=forum1("id")%>">
<U><B><%=forum1("grp")%></B></U></A>
<%'BU KATTA KAÇ KÝÞÝ VAR
sor = "SELECT * FROM online where altgrp="& forum1("id") &"  "
efkan.open sor, sur, 1, 3
adet=efkan.recordcount
efkan.Close
If adet=0 Then
else
Response.Write "<IMG SRC=""images/tekil.gif"" WIDTH=""17""  BORDER=""0"">&nbsp;"
Response.Write adet
End If
%>
<BR><%=forum1("aciklama")%>

<BR>
<!-- SUB DENEME -->
<%sor = "Select * from grup1 where pid="&forum1("id")&" order  by pidgrp asc"  
forum3.Open sor,forumbag,1,3
do while not forum3.eof  %>
<IMG SRC="images/file.gif" WIDTH="13" HEIGHT="13" BORDER="0" ALT="">
<A HREF="?part=altgrp&id=<%=forum("id")%>&pid=<%=forum1("id")%>&pid1=<%=forum3("id")%>">
<%=forum3("pidgrp")%></a>
<%forum3.movenext 
loop 
forum3.close
%>

</td>



<td width="5%" class="tdbrd" bgcolor="<%=bgcolor1%>" align="center" valign="center">
<%'SORU ADETÝ
sor = "Select altgrp from sorular where onay=1 and altgrp="&forum1("id")&"   "
forum2.Open sor,forumbag,1,3
adet=forum2.recordcount
forum2.close
Response.Write adet
%>
</td>

<td width="5%" class="tdbrd"  bgcolor="<%=bgcolor2%>"  align="center" valign="center">
<%'CEVAP ADETÝ
sor = "Select altgrp  from cevaplar where onay=1 and altgrp="&forum1("id")&"  "  
forum2.Open sor,forumbag,1,3
adet=forum2.recordcount
forum2.close
Response.Write adet
%>
</td>

<td width="30%" class="tdbrd"  bgcolor="<%=bgcolor1%>" align="left" valign="center">
<%'SON CEVAP
sor = "Select altgrp,uyeid ,kadi ,tarih,baslik,soruid,onay from cevaplar where onay=1 and altgrp="&forum1("id")&" order by id desc "  
forum2.Open sor,forumbag,1,3
If forum2.eof Then
Response.Write "<CENTER>-</CENTER>" 
Else 
'SON MESAJI VEREN ONLÝNE OLUP OLMADIÐI
sor="SELECT id,sontarih FROM uyeler WHERE id ="&forum2("uyeid")&"  "
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
efkan1.close %>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=forum2("uyeid")%>">
<%=forum2("kadi")%></A>&nbsp;<I><%=forum2("tarih")%></I>
<BR>
<!-- SON MESAJ OKU  LÝNKÝ -->
<div align="right">
<A HREF="?part=oku&id=<%=forum("id")%>&pid=<%=forum1("id")%>&urun=<%=forum2("soruid")%>">
<%=Left(forum2("baslik"),40)%>&nbsp;
<IMG SRC="forumimg/ok.gif" WIDTH="12" HEIGHT="12" BORDER="0" ALT=""></A></div>
<%End If
forum2.close%>
</td>
</tr>


<% 
forum1.movenext 
loop 
forum1.close
Response.Write "</table></td></tr>"
forum.movenext 
loop 
forum.close
%>
</table>


<%
set forum =Nothing
set forum1 =nothing
set forum2 =nothing
%>

