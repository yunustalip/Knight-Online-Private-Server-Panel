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


<div align="center">
<!--#INCLUDE file="forumayar.asp"-->
<!--#INCLUDE file="grpmenu.asp"-->

<%
id= kontrol(temizle(request.querystring("id")))
pid= kontrol(temizle(request.querystring("pid")))
pid1= kontrol(temizle(request.querystring("pid1")))
%>

<div align="right">
<A HREF="?part=yaz&gorev=soru&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>">
<IMG SRC="forumimg/yenikonu.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="">
</A></div>

<%
If pid1<>"" then
sor = "SELECT  * FROM  sorular WHERE altgrp="&pid&" and sub="&pid1&" and onay=1 order by id desc "  
Else
sor = "SELECT  * FROM  sorular WHERE  altgrp="&pid&" and onay=1 order by id desc "  
End If

forum.Open sor,forumbag,1,3
adet=forum.recordcount

if forum.eof or forum.bof then
Response.Write "<center><BR><BR><BR><IMG SRC=images/hata.gif WIDTH=19 BORDER=0 ALT=><BR>"
Response.Write "<B>Bu kategoride kayýt bulunamadý.. </B><P><a href=""javascript:history.back(1)""><B>&lt;&lt;Geri git</B></a>"
Response.End
End If

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If
%>


<!--ALT KATIN KONULARI DÖK-->
<table width="100%" bgcolor="" bordercolor="#CCFFFF"  border="0" cellspacing="0" cellpadding="4">
<tr>
<td background="forumimg/mn.gif" width="3%" height="38" align="center" valign="center">
<IMG SRC="forumimg/haber.gif" WIDTH="24" HEIGHT="24" BORDER="0" ALT="">
</td>

<td background="forumimg/mn.gif" width="35%" height="38" align="left" valign="center">
<FONT COLOR="#FFFFFF"><B><%=grupadi%> Konularý</B></FONT>
<%'BU KATTA KAÇ KÝÞÝ VAR
sor = "SELECT * FROM online where altgrp="& pid &"  "
efkan.open sor, sur, 1, 3
adet=efkan.recordcount
efkan.Close
Response.Write "<FONT COLOR=""#FFFFFF"">"
Response.Write "Aktif<IMG SRC=""images/tekil.gif"" WIDTH=""17""  BORDER=""0"">&nbsp;"
Response.Write adet
Response.Write "</FONT>"
%>
</td>

<td background="forumimg/mn.gif" width="15%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Konuyu Açan</B></FONT></td>

<td background="forumimg/mn.gif" width="2%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Okunma</B></FONT></td>

<td background="forumimg/mn.gif" width="2%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Bugun</B></FONT></td>

<td background="forumimg/mn.gif" width="2%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Mesaj</B></FONT></td>

<td background="forumimg/mn.gif" width="25%" height="38" align="center" valign="center">
<FONT COLOR="#FFFFFF"><B>Son Mesaj</B></FONT></td>
</tr>



<!-- SUB DENEME -->
<%sor = "Select * from grup1 where pid="&pid&" order  by pidgrp asc"  
forum3.Open sor,forumbag,1,3
If forum3.eof Then
Else %>
<tr><td colspan="7"  bgcolor="<%=bgcolor1%>" width="100%" height="38" align="center" valign="center">
<% do while not forum3.eof  %>
<IMG SRC="images/file.gif" WIDTH="13" HEIGHT="13" BORDER="0" ALT="">
<A HREF="?part=altgrp&id=<%=id%>&pid=<%=pid%>&pid1=<%=forum3("id")%>">
<%=forum3("pidgrp")%></a>
<%forum3.movenext 
loop %>
</td></tr>
<% forum3.close
End If %>






<% 
forum.pagesize =20
forum.absolutepage = shf
sayfa = forum.pagecount
for i=1 to forum.pagesize
if forum.eof then exit for
%>

<tr><td class="tdbrd" bgcolor="<%=bgcolor1%>" align="center" valign="center">
<%if forum("acik")=0 Then
Response.Write "<IMG SRC=""images/hata.gif"" WIDTH=""19""  BORDER=""0"">"
Else
Response.Write "<IMG SRC=""forumimg/yazi.gif"" WIDTH=""16""  BORDER=""0"">"
End If %>
</td>


<!-- BAÞLIK -->
<td class="tdbrd" bgcolor="<%=bgcolor2%>" align="left" valign="center">
<A HREF="?part=oku&id=<%=forum("grp")%>&pid=<%=forum("altgrp")%>&pid1=<%=forum("sub")%>&urun=<%=forum("id")%>">
<%=kucukharf(forum("baslik"))%></a>

<%'BU KATTA KAÇ KÝÞÝ VAR
sor = "SELECT * FROM online where urun="& forum("id") &"  "
efkan.open sor, sur, 1, 3
adet=efkan.recordcount
efkan.Close
If adet=0 Then
else
Response.Write "<IMG SRC=""images/tekil.gif"" WIDTH=""17""  BORDER=""0"">&nbsp;"
Response.Write adet
End If
%>
</td>

<!-- SORAN TARÝH -->
<td class="tdbrd" bgcolor="<%=bgcolor1%>"  align="left" valign="center">
<% 'SORAN VEREN ONLÝNE OLUP OLMADIÐI
sor="SELECT id,sontarih FROM uyeler WHERE id ="&forum("uyeid")&"  "
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
efkan1.close%>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=forum("uyeid")%>">
<%=forum("kadi")%></A><BR><%=forum("tarih")%>
</td>


<!-- HÝT -->
<td class="tdbrd" bgcolor="<%=bgcolor2%>" align="center" valign="center"><%=forum("hit")%></td>

<!-- GÜN HÝTÝ -->
<td class="tdbrd" bgcolor="<%=bgcolor1%>" align="center" valign="center">
<% tarih=day(date)&"."&month(date)&"."&year(date)
if forum("gun")=tarih then Response.Write  forum("gunhit") else Response.Write  0 End If %>
</td>

<!-- CEVAPLAR -->
<td class="tdbrd" bgcolor="<%=bgcolor2%>" align="center" valign="center">
<% sor = "Select  soruid from cevaplar where onay=1 and soruid="&forum("id")&" order by id desc"  
forum2.Open sor,forumbag,1,3
adet=forum2.recordcount
forum2.close
Response.Write adet
%></td>

<!-- EN SON CEVAP -->
<td class="tdbrd" bgcolor="<%=bgcolor1%>" align="left" valign="center">
<%
sor = "Select altgrp,uyeid ,kadi ,tarih,baslik,soruid from cevaplar where onay=1 and soruid="&forum("id")&" order by id desc " 
forum2.Open sor,forumbag,1,3
If forum2.eof then
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
efkan1.close
%>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=forum2("uyeid")%>">
<%=forum2("kadi")%></A>&nbsp;<I><%=forum2("tarih")%></I>
<BR>
<!-- SON MESAJ OKU  LÝNKÝ -->
<div align="right">
<A HREF="?part=oku&id=<%=id%>&pid=<%=pid%>&pid1=<%=forum("sub")%>&urun=<%=forum2("soruid")%>">
<%=kucukharf(Left(forum2("baslik"),40))%>&nbsp;
<IMG SRC="forumimg/ok.gif" WIDTH="12" HEIGHT="12" BORDER="0" ALT=""></A></div>
<%End If
forum2.close%>
</td>
</tr>


<%
forum.movenext
Next
forum.close %>
</table>
<!--DÖK SON -->


<P>
Sayfalar :
<% say=0
for y=1 to sayfa 
if say mod 10 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf ) then 'bulunduðun sayfaya link yok
Response.Write "<B>["&y&"]</B>"
else

Response.Write "<a href='default.asp?part=altgrp&id="&id&"&pid="&pid&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
Next

set forum =nothing
set forum1 =nothing
set forum2 =nothing

%>
<P>

