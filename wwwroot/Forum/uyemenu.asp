

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

<%
Session.LCID = 1055
Session.CodePage = 1254

'ONL�NE Z�YARET�� VE K���LER�N NERDE OLDU�UNU KAYDET
ipno=Request.ServerVariables("REMOTE_ADDR")
sor = "SELECT * FROM online where ipno='"& ipno &"'"
efkan.open sor, sur, 1, 3
If efkan.eof Then
efkan.addnew
efkan("ipno")=ipno
efkan("zaman")=now
efkan("uyeid")=Session("uyeid")
efkan("grp")  =kontrol(request.querystring("id"))
efkan("altgrp")=kontrol(request.querystring("pid"))
efkan("urun")=kontrol(request.querystring("urun"))
else
efkan("zaman")=Now
efkan("uyeid")=Session("uyeid")
efkan("grp")  =kontrol(request.querystring("id"))
efkan("altgrp")=kontrol(request.querystring("pid"))
efkan("urun")=kontrol(request.querystring("urun"))
End If
efkan.update
efkan.Close
%>

<div align="center">

<% hucre= " bgcolor="&bgcolor2&" align=""center""" %>
<%If Session("uyelogin")=True then%> 

<!-- �YE G�R�� YAPTIYSA �YE MENUSU  -->
<table  width="100%"  bgcolor="" bordercolor="#EAEAEA" border="0" cellspacing="4" cellpadding="0">
<tr height="16">
<%
'GELEN KUTUSU
sor="select kime , kimesildi  from mesaj where kime="&Session("uyeid")&" and kimesildi=0  "
efkan.Open sor,Sur,1,3
gelen=efkan.recordcount  
efkan.close

'OKUNMAYAN MESAJI VARMI 
sor="select  kime,okundu,kimesildi from mesaj where kime="&Session("uyeid")&" and okundu = 0 and kimesildi <>1 "
efkan.Open sor,Sur,1,3
okunmayan = efkan.recordcount
efkan.close

'G�DEN KUTUSU
sor="select kimden ,kimdensildi  from mesaj where kimden="&Session("uyeid")&" and kimdensildi=0  "
efkan.Open sor,Sur,1,3
giden=efkan.recordcount  
efkan.close
%>

<td <%=hucre%>><A HREF="?part=uyemesaj&gorev=gelen">Gelen (<%=gelen%>)</A></td>

<td <%=hucre%>>
<A HREF="?part=uyemesaj&gorev=gelen">
Okunmayan
<%If okunmayan > 0 Then 
Response.Write "<IMG SRC=""images/yanson1.gif"" WIDTH=""14"" BORDER=""0"">"
End If %>(<%=okunmayan%>)</A></td>

<td <%=hucre%>><A HREF="?part=uyemesaj&gorev=giden">Giden (<%=giden%>)</A></td>
<td <%=hucre%>><A HREF="default.asp?part=uyemesaj&gorev=yaz">Mesaj G�nder</A></td>
<td <%=hucre%>><A HREF="default.asp?part=uyegorev&gorev=bilgilerim">Bilgilerim</A></td>
<td <%=hucre%>><A HREF="default.asp?part=uyegorev&gorev=uyeler">�yeler</A></td>
<td <%=hucre%>><A HREF="default.asp?part=uyegorev&gorev=cikis">��k��</A></td></tr></table>

<!-- ADM�N G�R��� �SE -->
<%If Session("efkanlogin")=True then%>
<A onclick="javascript:toggle('aa');return false;" href="#"><B>Admin Menusu</B></A>
<table id="aa" style="DISPLAY: none" width="100%"  bgcolor="" bordercolor="#EAEAEA" border="0" cellspacing="5" cellpadding="0">
<tr>
<td <%=hucre%>><A HREF="?part=kat">Kategori Y�net</A></td>
<td <%=hucre%>><A HREF="?part=uyegorev&gorev=eskisil">5 Ayd�r Girmeyenleri Sil</A></td>
<td <%=hucre%>><A HREF="?part=uyegorev&gorev=aktifolmayansil">�yeliklerini Aktifle�tirmeyenleri Sil</A></td>
<td <%=hucre%>><A HREF="?part=uyegorev&gorev=mesajtemizle">3 ayl�k mesajlar� sil</A></td>
<td <%=hucre%>><A HREF="?part=uyegorev&gorev=yasakli">Yasakl�lar</A></td>

</tr><tr>

<td <%=hucre%>><A HREF="?part=uyegorev&gorev=yasakliekle">Yasakl� Ekle</A></td>
<td <%=hucre%>><A HREF="?part=onaysiz&gorev=onaysiz&neresi=sorular">Onays�z Konular</A></td>
<td <%=hucre%>><A HREF="?part=onaysiz&gorev=onaysiz&neresi=cevaplar">Onays�z Cevaplar</A></td>
<td <%=hucre%>><A HREF="?part=onayli&gorev=onayli&neresi=sorular">Onayl� Konular</A></td>
<td <%=hucre%>><A HREF="?part=onayli&gorev=onayli&neresi=cevaplar">Onayl� Cevaplar</A></td>

</tr><tr>
<td <%=hucre%>><A HREF="?part=uyegorev&gorev=aktivasyonsuzlar">�y.Aktif.Emali G�nder</A></td>
<td <%=hucre%>><A HREF="?part=sayac">Saya�</A></td>
</tr>
</table>
<%
End If

Else
%>
<!-- �YE G�R��� YOKSA �YE G�R��� MENUSU -->
<form action="uyegorev.asp?gorev=kontrol" method="POST" >
<table  width="100%"  border="0" cellspacing="0" cellpadding="0">
<tr><td class="tdbrd"  height="30" width="100%" align="left" valign="center">
&nbsp;Kullan�c� Ad� <input name="kadi" size="15" maxlength="20">
�ifreniz <input type="password" name="sifre" size="15" maxlength="20">
<input type="submit" value="�ye Giri�i">
&nbsp;
<A HREF="default.asp?part=uyegorev&gorev=uyeol">�ye Ol</A> | 
<A HREF="default.asp?part=uyegorev&gorev=unuttum">Unuttum</A> | 
<A HREF="default.asp?part=uyegorev&gorev=emailgelmedi">Aktivasyon Emaili Gelmedi</A>
</td></tr>
</form>
</table>
<%End If%>                   