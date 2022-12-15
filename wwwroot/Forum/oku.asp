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
'Response.Expires=0
Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage
%>

<!--#INCLUDE file="forumayar.asp"-->
<!--#INCLUDE file="grpmenu.asp"-->
<P>
<div align="center">

<% 
Response.Buffer = True 

id= kontrol(temizle(request.querystring("id")))
pid= kontrol(temizle(request.querystring("pid")))
pid1= kontrol(temizle(request.querystring("pid1")))
urun= kontrol(temizle(request.querystring("urun")))

sor  = "select * from sorular where id="&urun&""
forum.Open sor,forumbag,1,3
acik=forum("acik")

abcd =  Request.ServerVariables("REMOTE_ADDR") 
if  Session ("abcd")=abcd and Session ("urun") = urun   then
else

tarih=day(date)&"."&month(date)&"."&year(date)
if trim(forum("gun")) <> tarih or trim(forum("gun"))=""  then
forum("hit") = forum("hit") + 1
forum("gun") = tarih
forum("gunhit") = 1
forum.Update
else
forum("hit") = forum("hit") +1
forum("gunhit") = forum("gunhit") + 1
forum.Update
End If

Session ("abcd") =Request.ServerVariables("REMOTE_ADDR") 
Session ("urun") = urun
End If 
%>

<a name="top">

<!-- SORU YAZ CEVAP YAZ BUTON FORUM AÇIKSA -->
<div align="right">
<%If acik=1 Then %>
<A HREF="?part=yaz&gorev=cevap&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>&urun=<%=urun%>">
<IMG SRC="forumimg/cevapla.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<A HREF="?part=oku&gorev=takip&id=<%=id%>&pid=<%=pid%>&urun=<%=urun%>">
<IMG SRC="forumimg/takip.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<%else%>
<IMG SRC="images/hata.gif" WIDTH="19" HEIGHT="19" BORDER="0" ALT="">Bu konu kapatýldý...
<%End If%>
<A HREF="?part=yaz&gorev=soru&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>">
<IMG SRC="forumimg/yenikonu.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
</div>


<!-- SORU METNÝ -->
<table width="100%" bgcolor="" bordercolor="<%=bgcolor1%>" border="0" cellspacing="0" cellpadding="4">
<tr height="38">
<td background="forumimg/mn.gif" width="15%"  align="left" valign="center">
<FONT COLOR="#FFFFFF"><B>Konuyu Açan</B></FONT></td>
<td background="forumimg/mn.gif" width="70%" align="left" valign="center">
<FONT COLOR="#FFFFFF"><B>Konu</B></FONT>
&nbsp;&nbsp;
<%'BU KATTA KAÇ KÝÞÝ VAR
sor = "SELECT * FROM online where urun="& urun &"  "
efkan.open sor, sur, 1, 3
adet=efkan.recordcount
efkan.Close
Response.Write "<FONT COLOR=""#FFFFFF"">Aktif <IMG SRC=""images/tekil.gif"" WIDTH=""17""  BORDER=""0"">&nbsp;"
Response.Write adet &"</FONT>"
%>
</td>

<td background="forumimg/mn.gif" width="25%" align="left" valign="center"><!-- ÝÞLEM --></td>
</tr>


<tr>
<td bgcolor="<%=bgcolor2%>" align="left" valign="top">

<%sor="SELECT * FROM uyeler WHERE id ="&forum("uyeid")&"  "
efkan1.Open sor,Sur,1,3

If efkan1.eof Then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
Response.Write "&nbsp;<B>" &forum("kadi") & "</B>"
silinenuye=1
Else
zaman=datediff("n",efkan1("sontarih"),Now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>" 
End If
%>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=efkan1("id")%>">
<B><%=efkan1("kadi")%></B></A>
<% End If %>

</td>

<td  bgcolor="<%=bgcolor1%>" align="left" valign="center">
<B>Eklenme/Güncelleme Trh :</B> <%=forum("tarih")%>&nbsp;
<B>Toplam Okunma :</B> <%=forum("hit")%>&nbsp;
<B>Bugün  :</B>
<% tarih=day(date)&"."&month(date)&"."&year(date)
if forum("gun")=tarih then Response.Write  forum("gunhit") else Response.Write  0 End If %>
</td>


<td bgcolor="<%=bgcolor1%>" align="right">
<!-- DEÐÝÞTÝRME BUTONU -->
<%If acik=1 And session("uyeid")=forum("uyeid") Or Session("efkanlogin")=True Then %>
<A HREF="?part=degistir&gorev=degistir&yer=soru&id=<%=forum("id")%>">
<IMG SRC="forumimg/degistir.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="">
</A>
<%End If %>
<!-- AÇIKSA ALINTI YAP -->
<%If acik=1 Then %>

<A HREF="?part=yaz&gorev=cevap&yer=soru&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>&urun=<%=urun%>">
<IMG SRC="forumimg/alinti.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<%End If%>

</td>
</tr>


<!--  -->
<tr>
<td bgcolor="<%=bgcolor1%>" align="left" valign="top">
<% If silinenuye<>1 Then 

sor = "Select * from uyeresim where uyeid = "&efkan1("id")&" " 
efkan2.Open sor,Sur,1,3
if efkan2.eof or efkan2.bof Then%>
<IMG SRC="avatar/<%=efkan1("avatar")%>" WIDTH="130"  BORDER="0" ALT=""><BR>
<% Else %>
<IMG SRC="uyeler/<%=efkan2("uyeresim")%>" WIDTH="130"  BORDER="0" ALT=""><BR>
<% End If
efkan2.close



birim=hitbirim
kat=efkan1("hit")/ birim
kat = Fix(kat)
if kat=0 then 
Response.Write "<IMG SRC=""images/yildiz.gif"" WIDTH=""16"" BORDER=""0"" alt=""toplam"&efkan1("hit")&"giriþ"">"
else
for sira=1 to kat 
Response.Write "<IMG SRC=""images/yildiz.gif"" WIDTH=""16"" BORDER=""0"" alt=""toplam"&efkan1("hit")&"giriþ"">"
next 
End If
%> 
<BR>
<B>Üye No:&nbsp;</B><%=efkan1("id")%><BR>
<B>Üye Oluþu:</B><BR><%=efkan1("tarih")%><BR>
<B>Son Giriþ:</B><BR><%=efkan1("sontarih")%><BR>
<B>Nerden:</B><BR><%=efkan1("sehir")%><BR>


<%'BU UYENÝN MESAJ VE KONU SAYISI
sor = "Select * from sorular where onay=1 and  uyeid="&forum("uyeid")&" "  
forum2.Open sor,forumbag,1,3
soruadet=forum2.recordcount
forum2.close
sor = "Select * from cevaplar where onay=1 and  uyeid="&forum("uyeid")&" "  
forum2.Open sor,forumbag,1,3
cevapadet=forum2.recordcount
forum2.close
Response.Write "<B>Konu  &nbsp;:</B>&nbsp;" &soruadet &"<BR><B>Mesaj :</B>&nbsp;"&cevapadet

Else 
Response.Write "<B>Bu üye silindi</B>"
End If

%>

</td>

<td bgcolor="<%=bgcolor2%>" colspan="2" align="left" valign="top">
<B><%=kucukharf(forum("baslik"))%></B>
<P>
<%=forum("aciklama")%>
<P>

<% If silinenuye=1 Then 
else
%>
------------------------------------------------------------------------------------------<BR>
<%=editor(efkan1("imza"))%>

<P>
<% if efkan1("url")="http://" Or efkan1("url")="" Then
else%>
<A HREF="<%=efkan1("url")%>" target="_blank">
<IMG SRC="forumimg/web.gif" WIDTH="68" HEIGHT="17" BORDER="0" ALT="<%=efkan1("url")%>"></A>
<% End If%>

<A HREF="default.asp?part=uyemesaj&gorev=yaz&id=<%=efkan1("id")%>&kime=<%=efkan1("kadi")%>">
<IMG SRC="forumimg/umsg.gif" WIDTH="68" HEIGHT="17" BORDER="0" ALT="Bu üyeye mesaj yaz"></A>


<%
End If
efkan1.close

silinenuye =0

'ADMÝN ÝSE SORU SÝL BUTONU
If Session("efkanlogin")=True  Then %>
<DIV ALIGN="RIGHT">
<A HREF="?part=gorev&gorev=sorusil&id=<%=id%>&pid=<%=pid%>&urun=<%=urun%>">
<IMG SRC="forumimg/ksil.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>

<A HREF="?part=gorev&gorev=tasi&id=<%=urun%>">
<IMG SRC="forumimg/tasi.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>

<A HREF="?part=gorev&gorev=subtasi&pid=<%=pid%>&urun=<%=urun%>">
<IMG SRC="forumimg/sub.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>

<%If acik=1 Then %>
<A HREF="?part=gorev&gorev=kapa&id=<%=id%>&pid=<%=pid%>&urun=<%=urun%>">
<IMG SRC="forumimg/kapat.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<% Else %>
<A HREF="?part=gorev&gorev=ac&id=<%=id%>&pid=<%=pid%>&urun=<%=urun%>">
<IMG SRC="forumimg/ac.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<% End If %>
</DIV>
<%End If%>


</td></tr>
</TABLE>



<!-- SORU YAZ CEVAP YAZ BUTON -->
<div align="right">
<%If acik=1 Then %>
<A HREF="?part=yaz&gorev=cevap&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>&urun=<%=urun%>">
<IMG SRC="forumimg/cevapla.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<%End If%>
<A HREF="?part=yaz&gorev=soru&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>">
<IMG SRC="forumimg/yenikonu.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
</div>

<BR>



<!-- CEVAPLAR -->
<%
soruid=forum("id")
adi=forum("baslik")
forum.close

sor = "Select * from cevaplar where soruid = "&soruid&" and onay=1 order by id asc " 
forum.Open sor,forumbag,1,3
adet=forum.recordcount
if forum.eof or forum.bof then
Response.Write "<B>Bu konu hakkýnda cevap yazýlmadý...</B>"
Else

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If
%>

<B>Bu konuya <%=adet%> adet cevap yazýlmýþ</B>
<table width="100%" bgcolor="" bordercolor="<%=bgcolor1%>" border="0" cellspacing="0" cellpadding="4">
<tr>
<td background="forumimg/mn.gif" width="15%" height="38" align="left" valign="center">
<FONT COLOR="#FFFFFF"><B>Cevaplayanlar</B></FONT></td>
<td background="forumimg/mn.gif" width="50%" height="38" align="left" valign="center">
<FONT COLOR="#FFFFFF"><B>Mesaj</B></FONT></td>
<td background="forumimg/mn.gif" width="35%" align="left" valign="center"><!-- ÝÞLEM --></td>
</tr>
<% 
forum.pagesize =25 '1 sayfada görüntülemek istediðiniz kayýt sayýsý (deðiþtirebilirsiniz)
forum.absolutepage = shf
sayfa = forum.pagecount
for i=1 to forum.pagesize
if forum.eof then exit for
%>

<tr>
<td bgcolor="<%=bgcolor2%>" align="left" valign="center">

<%sor="SELECT * FROM uyeler WHERE id ="&forum("uyeid")&"  "
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
Response.Write "&nbsp;<B>" &forum("kadi") & "</B>"
silinenuye=1
else
zaman=datediff("n",efkan1("sontarih"),Now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>" 
End If
%>
<A HREF="?part=uyegorev&gorev=uyebilgi&id=<%=efkan1("id")%>">
<B><%=efkan1("kadi")%></B></A>
<% End If %>
</td>

<td bgcolor="<%=bgcolor1%>" align="left" valign="center">
<B>Ýd:</B><%=forum("id")%>&nbsp; <B>Eklenme/Güncelleme Trh :</B> <%=forum("tarih")%>
</td>


<td bgcolor="<%=bgcolor1%>" align="right" valign="center">
<!-- DEÐÝÞTÝRME BUTONU -->
<%If acik=1 And session("uyeid")=forum("uyeid") Or Session("efkanlogin")=True Then %>
<A HREF="?part=degistir&gorev=degistir&yer=cevap&id=<%=forum("id")%>">
<IMG SRC="forumimg/degistir.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<%End If %>
<!-- AÇIKSA ALINTI YAP -->
<%If acik=1 Then %>

<A HREF="?part=yaz&gorev=cevap&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>&urun=<%=urun%>">
<IMG SRC="forumimg/cevapla.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<A HREF="?part=yaz&gorev=cevap&yer=cevap&id=<%=id%>&pid=<%=pid%>&pid1=<%=pid1%>&urun=<%=urun%>&alid=<%=forum("id")%>">
<IMG SRC="forumimg/alinti.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
<%End If%>
<a href="#top" title="Konuya Git"><IMG SRC="images/up.gif" WIDTH="20" HEIGHT="20" BORDER="0" ALT=""></a>
</td>

</tr>

<tr>
<td bgcolor="<%=bgcolor1%>" align="left" valign="top">
<% If silinenuye<>1 Then 

sor = "Select * from uyeresim where uyeid = "&efkan1("id")&" " 
efkan2.Open sor,Sur,1,3
if efkan2.eof or efkan2.bof Then%>
<IMG SRC="avatar/<%=efkan1("avatar")%>" WIDTH="130"  BORDER="0" ALT=""><BR>
<% Else %>
<IMG SRC="uyeler/<%=efkan2("uyeresim")%>" WIDTH="130"  BORDER="0" ALT=""><BR>
<% End If
efkan2.close


birim=hitbirim
kat=efkan1("hit")/ birim
kat = Fix(kat)
if kat=0 then 
Response.Write "<IMG SRC=""images/yildiz.gif"" WIDTH=""16"" BORDER=""0"" alt=""toplam"&efkan1("hit")&"giriþ"">"
else
for sira=1 to kat 
Response.Write "<IMG SRC=""images/yildiz.gif"" WIDTH=""16"" BORDER=""0"" alt=""toplam"&efkan1("hit")&"giriþ"">"
next 
End If
%> 
<BR>
<B>Üye No:</B>&nbsp;<%=efkan1("id")%><BR>
<B>Üye Oluþu:</B><BR><%=efkan1("tarih")%><BR>
<B>Son Giriþ:</B><BR><%=efkan1("sontarih")%><BR>
<B>Nerden:</B><BR><%=efkan1("sehir")%><BR>


<%'BU UYENÝN MESAJ VE KONU SAYISI
sor = "Select * from sorular where onay=1 and  uyeid="&forum("uyeid")&" "  
forum2.Open sor,forumbag,1,3
soruadet=forum2.recordcount
forum2.close
sor = "Select * from cevaplar where onay=1 and uyeid="&forum("uyeid")&" "  
forum2.Open sor,forumbag,1,3
cevapadet=forum2.recordcount
forum2.close
Response.Write "<B>Konu  &nbsp;:</B>&nbsp;" &soruadet &"<BR><B>Mesaj :</B>&nbsp;"&cevapadet

Else 
Response.Write "<B>Bu üye silindi</B>"
End If
%>

</td>

<td bgcolor="<%=bgcolor2%>" colspan="2" align="left" valign="top">
<B><%=kucukharf(forum("baslik"))%></B>
<P>
<%=forum("alinti")%><P>
<%=forum("aciklama")%>
<P>

<% If silinenuye=1 Then 
else
%>
------------------------------------------------------------------------------------------<BR>
<%=editor(efkan1("imza"))%>

<P>
<% if efkan1("url")="http://" Or efkan1("url")="" Then
else%>
<A HREF="<%=efkan1("url")%>" target="_blank">
<IMG SRC="forumimg/web.gif" WIDTH="68" HEIGHT="17" BORDER="0" ALT="<%=efkan1("url")%>"></A>
<% End If%>

<A HREF="default.asp?part=uyemesaj&gorev=yaz&id=<%=efkan1("id")%>&kime=<%=efkan1("kadi")%>">
<IMG SRC="forumimg/umsg.gif" WIDTH="68" HEIGHT="17" BORDER="0" ALT="Bu üyeye mesaj yaz"></A>


<%
End If
efkan1.close

silinenuye=0

'ADMÝN ÝSE SORU SÝL BUTONU
If Session("efkanlogin")=True  Then %>
<DIV ALIGN="RIGHT">
<A HREF="?part=gorev&gorev=cevapsil&id=<%=id%>&pid=<%=pid%>&urun=<%=urun%>&cevapid=<%=forum("id")%>">
<IMG SRC="forumimg/msil.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT=""></A>
</DIV>
<%End If%>

</td></tr>


<%
forum.movenext
Next
forum.close
%>

</TABLE>


<%If acik<>1 Then %>
<IMG SRC="images/hata.gif" WIDTH="19" HEIGHT="19" BORDER="0" ALT="">Bu konu kapatýldý...
<%End If%>
<P>

Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 10 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf ) then 
Response.Write "<B>["&y&"]</B>"
else
Response.Write "<a href='default.asp?part=oku&id="&id&"&pid="&pid&"&urun="&urun&"&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
next


End If



'KONU TAKÝP LÝSTESÝ
gorev= temizle(request.querystring("gorev"))
If gorev="takip" Then 
Response.Buffer = True 

If Session("uyelogin")=True <> True Then 
Response.Redirect ("?part=uyegorev&gorev=girisform")
Response.End
End If

sor="SELECT * FROM takip where urun="&urun&" and uyeid="&session("uyeid")&"  "
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
efkan1.close
sor="SELECT * FROM takip  "
efkan1.Open sor,Sur,1,3
efkan1.AddNew
efkan1("grp") = id
efkan1("altgrp") =pid
efkan1("urun") =urun
efkan1("kadi") =session("kadi")
efkan1("uyeid") =session("uyeid")
efkan1.Update
efkan1.close
Else
efkan1.close
End If







End If






set forum =Nothing
set forum1 =Nothing
set forum2 =Nothing

%>



<P>







