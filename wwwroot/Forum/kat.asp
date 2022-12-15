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

<%

Response.Buffer = True


If  Session("efkanlogin")<>True  Then
Response.Write "Bu iþlem için yetkiniz yok"
Response.End
End If 


gorev      =temizle(Request.QueryString("gorev"))
islem     =temizle(Request.QueryString("islem"))
id         =kontrol(Request.QueryString("id"))


'///////////////////////////  TÜM KATEGORÝLER ////////////////////////
If gorev="" Or gorev="map" Then %>
<A HREF="?part=kat&gorev=map"><B>KATEGORÝ YÖNETÝMÝ ANA SAYFASI</B></A>
<P>

<A HREF="?part=kat&gorev=map&islem=grupekle"><B>Yeni Ana Kategori Ekle</B></A><P>

<% If islem="grupekle" Then %>
<form method="POST" action="?part=kat&gorev=map&islem=grupekle">
<B>ANA KATEGORÝ EKLÝYORSUNUZ</B>
<table width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Yeni Ana Kategori Adý</td>
<td width="70%">*<input name="grp" type="text" value="" size="60"></td></tr>
<tr><td>Açýklama</td>
<td>*<input name="aciklama" type="text" value="" size="60"></td></tr>
<tr><td colspan="2" align=	"center">
<input type="submit" value=" Ekle " name="submit" >&nbsp;<INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<%
if request.form("grp")="" Then 
Else
sor = "Select * from grup where grp = '" & Trim(request.form("grp")) & "'  " 
forum.Open sor,forumbag,1,3
adet=forum.recordcount
if adet >0  then
Response.Write "<script language='JavaScript'>alert('Bu kategori zaten var...');</script>"
forum.close
else
forum.AddNew
forum("grp")               = Trim (Request.Form ("grp"))
forum("aciklama")        = Trim (Request.Form ("aciklama"))
forum.Update
forum.close
Response.Redirect "?part=kat&gorev=map"
End If
End If 
End If


If islem="grupdegistir" Then
sor  = "select * from grup where id="&id&""
forum.Open sor,forumbag,1,3 %>
<P>
<FONT COLOR="blue">Bu kategori adý deðiþir ve bu kategoriye ait alt kategori ve veriler bu kategoriye baðlý kalmaya devam ederler...</FONT>
<form method="POST" action="?part=kat&gorev=map&islem=grupdegistir&id=<%=id%>">
<B>ANA KATEGORÝ ADINI DEÐÝÞTÝRÝYORSUNUZ</B>
<table width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Ana Kategori Adý</td>
<td width="70%">*<input name="grp" type="text" value="<%=forum("grp")%>" size="60"></td></tr>
<tr><td>Açýklama</td>
<td><input name="aciklama" type="text" value="<%=forum("aciklama")%>" size="60"></td></tr>
<tr><td colspan="2" align=	"center">
<input type="submit" value=" Deðiþtir " name="submit" >&nbsp;<INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<%
if request.form("grp")="" then
forum.close
Else
forum("grp")               = Trim (Request.Form ("grp"))
forum("aciklama")        = Trim (Request.Form ("aciklama"))
forum.Update
forum.close
Response.Redirect "?part=kat&gorev=map"
End If
End If 


If islem="altgrupekle" Then %>
<form method="POST" action="?part=kat&gorev=map&islem=altgrupekle&id=<%=id%>">
<B>ALT KATEGORÝ EKLÝYORSUNUZ</B>
<table width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Alt Kategori Adý</td>
<td width="70%">*<input name="grp" type="text" value="" size="60"></td></tr>
<tr><td>Açýklama</td>
<td><input name="aciklama" type="text" value="" size="60"></td></tr>
<tr><td colspan="2" align=	"center">
<input type="submit" value=" Ekle " name="submit" >&nbsp;<INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<% 
if request.form("grp")="" Then 
Else
sor = "Select * from grup where grp=  '"&Trim(request.form("grp"))&"' and altgrp= "& id &" " 
forum.Open sor,forumbag,1,3
adet=forum.recordcount
if adet >0  then
Response.Write "<script language='JavaScript'>alert('Bu kategoride bu alt kategori zaten var...');</script>"
forum.close
else
forum.AddNew
forum("altgrp")            = id
forum("grp")                = Trim (Request.Form ("grp"))
forum("aciklama")        = Trim (Request.Form ("aciklama"))
forum.Update
forum.close
Response.Redirect "?part=kat&gorev=map"
End If
End If 
End If


If islem="altgrupdegistir" Then
sor  = "select * from grup where id="&id&""
forum.Open sor,forumbag,1,3 %>
<P>
<FONT COLOR="red">Bu alt kategori adý deðiþir  veriler bu kategoriye baðlý kalmaya devam ederler...</FONT>
<form method="POST" action="?part=kat&gorev=map&islem=altgrupdegistir&id=<%=id%>">
<B>ALT KATEGORÝ ADINI DEÐÝÞTÝRÝYORSUNUZ</B>
<table width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Alt Kategori Adý</td>
<td width="60%">*<input name="grp" type="text" value="<%=forum("grp")%>" size="50"></td></tr>
<tr><td>Açýklama</td>
<td><input name="aciklama" type="text" value="<%=forum("aciklama")%>" size="50"></td></tr>
<tr><td colspan="2" align=	"center">
<input type="submit" value=" Deðiþtir " name="submit" >&nbsp;<INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<%
if request.form("grp")="" then
forum.close
Else
forum("grp")                = Trim (Request.Form ("grp"))
forum("aciklama")        = Trim (Request.Form ("aciklama"))
forum.Update
forum.close
Response.Redirect "?part=kat&gorev=map"
End If
End If 

if islem="grupsil" Then 
sor = "UPDATE sorular SET onay = 0 WHERE grp="&id&" or altgrp="&id&" "
forum.open sor,forumbag,1,3
sor = "DELETE from grup WHERE id="&id&""
forum.open sor,forumbag,1,3
sor = "DELETE from grup WHERE altgrp="&id&""
forum1.open sor,forumbag,1,3
sor = "DELETE from grup1 WHERE pid="&id&""
forum2.open sor,forumbag,1,3
Response.Redirect "?part=kat&gorev=map"
End If



If islem="subdegistir" then
id=Request.QueryString("id")
sor  = "select * from grup1 where id="&id&" "
forum.open sor,forumbag,1,3
%>
<form method="POST" action="?part=kat&gorev=map&islem=subdegistir&id=<%=id%>">
<B>ALT SUB ADINI DEÐÝÞTÝRÝYORSUNUZ</B>
<table width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Alt Sub Adý </td>
<td width="60%">*<input name="pidgrp" type="text" value="<%=forum("pidgrp")%>" size="60"></td></tr>
<tr><td>Açýklama</td>
<td><input name="pidgrpaciklama" type="text" value="<%=forum("pidgrpaciklama")%>" size="60"></td></tr>
<tr><td colspan="2" align=	"center">
<input type="submit" value=" Deðiþtir " name="submit" >&nbsp;<INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<%
if request.form("pidgrp")="" Then
forum.close
Else
forum("pidgrp") = Trim (Request.Form ("pidgrp"))
forum("pidgrpaciklama") = Trim (Request.Form ("pidgrpaciklama"))
forum.Update
forum.close
Response.Redirect "?part=kat&gorev=map"
End If
End If

If islem="subsil" then
id=Request.QueryString("id")
sor = "DELETE from grup1 WHERE id="&id&""
forum.open sor,forumbag,1,3
Response.Redirect "?part=kat&gorev=map"
End If



If islem="subekle" then
id=Request.QueryString("id") %>
<form method="POST" action="?part=kat&gorev=map&islem=subekle&id=<%=id%>">
<B>ALT SUB EKLÝYORSUNUZ</B>
<table width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Alt Sub Adý </td>
<td width="60%">*<input name="pidgrp" type="text" value="" size="60"></td></tr>
<tr><td>Açýklama</td>
<td><input name="pidgrpaciklama" type="text" value="" size="60"></td></tr>
<tr><td colspan="2" align=	"center">
<input type="submit" value=" Ekle " name="submit" >&nbsp;<INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<%
if request.form("pidgrp")=""   then
Else
sor  = "select * from grup1 "
forum.open sor,forumbag,1,3
forum.AddNew
forum("pid") = Trim(id)
forum("pidgrp") = Trim (Request.Form ("pidgrp"))
forum("pidgrpaciklama") = Trim (Request.Form ("pidgrpaciklama"))
forum.Update
forum.close
Response.Redirect "?part=kat&gorev=map"
End If
End If




%>

<!-- KATEOLAR DÖKÜYORUM -->
<table width="90%" bgcolor="" bordercolor="<%=bgcolor1%>" border="1" cellspacing="0" cellpadding="3">
<tr><td colspan="4"><FONT COLOR="Blue">
<B>Kategori sildiðinizde kategori ve alt kategorileri silinir .
Silinen kategorilere ait kayýtlar silinmez kategorisiz kalýr ve onaysýz duruma geçerek yayýnlanmaz.
Kategorisiz olan bu kayýtlarý baþka kategorilere taþýyabilir veya yeni kategoriler oluþturup bu kategorilere taþýyabilirsiniz.</B></FONT>
</td></tr>

<tr bgcolor="">
<td width="70%"><B>Kategori ve Alt Kategoriler </B></td>
<td width="10%" align="center"><B>Deðiþtir</B></td>
<td width="10%"align="center"><B>Sil</B></td>
</tr>
<%sor = "Select * from grup where altgrp=0 order by grp asc"  
forum.Open sor,forumbag,1,3

if forum.eof or forum.bof then
Response.End
End If
do while not forum.eof  %>
<tr bgcolor="<%=bgcolor1%>"><td>
<IMG SRC="images/ok.gif" WIDTH="12" HEIGHT="12" BORDER="0" ALT="">
<A HREF="?part=grp&id=<%=forum("id")%>"><B><%=buyukharf(forum("grp"))%></B></A>
&nbsp;&nbsp;
<A HREF="?part=kat&gorev=map&islem=altgrupekle&id=<%=forum("id")%>"><B>+ Alt Kategori Ekle</B></A>
</td>

<td align="center">
<A HREF="?part=kat&gorev=map&islem=grupdegistir&id=<%=forum("id")%>">Deðiþtir</A></td>
<td align="center"><A HREF="?part=kat&gorev=map&islem=grupsil&id=<%=forum("id")%>" onClick="return submitConfirm(this)">Sil</A></td>
</tr>

<%sor = "Select * from grup where altgrp="&forum("id")&" order  by grp asc"  
forum1.Open sor,forumbag,1,3
do while not forum1.eof  %>

<tr><td>&nbsp;&nbsp;&nbsp;
<IMG SRC="images/kare.gif" WIDTH="10" HEIGHT="9" BORDER="0" ALT="">
<A HREF="?part=altgrp&id=<%=forum("id")%>&pid=<%=forum1("id")%>"><B><%=forum1("grp")%></B></A>
 &nbsp;&nbsp;&nbsp;
<A HREF="?part=kat&gorev=map&islem=subekle&id=<%=forum1("id")%>">+Alt Sub ekle</A>
</td>

<td align="center"><A HREF="?part=kat&gorev=map&islem=altgrupdegistir&id=<%=forum1("id")%>">Deðiþtir</A>
<td align="center"><A HREF="?part=kat&gorev=map&islem=grupsil&id=<%=forum1("id")%>" onClick="return submitConfirm(this)">Sil</A></td>
</tr>

<!-- SUB DÖK -->
<%sor = "Select * from grup1 where pid="&forum1("id")&" order  by pidgrp asc"  
forum2.Open sor,forumbag,1,3
do while not forum2.eof  %>
<tr><td><FONT COLOR="blue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<IMG SRC="images/ok2.jpg" WIDTH="14" HEIGHT="14" BORDER="0" ALT="">
<A HREF="?part=altgrp&id=<%=forum("id")%>&pid=<%=forum1("id")%>&pid1=<%=forum2("id")%>">
<%=forum2("pidgrp")%></FONT></a></td>
<td align="center">
<A HREF="?part=kat&gorev=map&islem=subdegistir&id=<%=forum2("id")%>">Deðiþtir</A></td>
<td align="center">
<A HREF="?part=kat&gorev=map&islem=subsil&id=<%=forum2("id")%>" onClick="return submitConfirm(this)">Sil</A></td></tr>
<%forum2.movenext 
loop 
forum2.close


forum1.movenext 
loop 
forum1.close
forum.movenext 
loop 
forum.close%>
</table>
<%
End If

%>




