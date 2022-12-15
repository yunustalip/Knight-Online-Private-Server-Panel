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
<!--#INCLUDE file="koru.asp"-->

<div align="center">

<SCRIPT language=JavaScript>
		function AddForm(form)
			{
				document.formcevap.yorum.value = document.formcevap.yorum.value + form
				document.formcevap.yorum.focus();
			}

function MesajKodu(Kod,TamYazi, YaziEkle) {
if (Kod != "") {
KodEkle = prompt(TamYazi + "\nÖrnek : http://www.makineteknik.com", YaziEkle);
if ((KodEkle != null) && (KodEkle != "")){
document.formcevap.yorum.value += "[" + Kod + "#" + KodEkle + "#]" + KodEkle + "[/"+ Kod + "] ";
}
}
document.formcevap.yorum.focus();
}

function email(Kod,TamYazi, YaziEkle) {
if (Kod != "") {
KodEkle = prompt(TamYazi + "\nÖrnek : info@makineteknik.com", YaziEkle);
if ((KodEkle != null) && (KodEkle != "")){
document.formcevap.yorum.value += "[" + Kod + "#" + KodEkle + "#]" + KodEkle + "[/"+ Kod + "] ";
}
}
document.formcevap.yorum.focus();
}

function resimekle(Kod,TamYazi, YaziEkle) {
if (Kod != "") {
KodEkle = prompt(TamYazi + "\nÖrnek : http://www.makineteknik.com/banner.gif", YaziEkle);
if ((KodEkle != null) && (KodEkle != "")){
document.formcevap.yorum.value += "[" + Kod + "#" + KodEkle + "#]";
}
}
document.formcevap.yorum.focus();
}

function FormDiger(Kod,TamYazi, YaziEkle) {
if (Kod != "") {
KodEkle = prompt(TamYazi + "\n", YaziEkle);
if ((KodEkle != null) && (KodEkle != "")){
document.formcevap.yorum.value += "[" + Kod + "]" + KodEkle + "[/"+ Kod + "] ";
}
}
document.formcevap.yorum.focus();
}
</SCRIPT>

<% 
Response.Buffer = True 
session.TimeOut = 600 
id=kontrol(Request.QueryString("id"))
yer=temizle(Request.QueryString("yer"))
gorev=temizle(Request.QueryString("gorev"))

If gorev="degistir" then
  If yer="soru" then
  sor = "select  * from  sorular where id="&id&"  "
  else
  sor = "select  * from cevaplar where id="&id&"  "
  End If
forum.Open sor,forumbag,1,3

If session("uyeid")=forum("uyeid") Or Session("efkanlogin")=True  Then
else
Response.Write "Bu iþlem için yetkiniz yok"
Response.End
End If 

session("yazdim")=false

%>
<BR><BR>
<form name="formcevap" onsubmit="return formCheck(this);" action="?part=degistir&gorev=kaydet&yer=<%=yer%>&id=<%=forum("id")%>" method="post" >

<table width="100%" border="1" bgcolor="" bordercolor="#FFFFFF" align="center" cellpadding="0" cellspacing="0">
<tr height=""><td width="100%" align="center" valign="center">


<B>Konu Baþlýðý :&nbsp;</B>
<INPUT TYPE="text" NAME="Baslik" size="100" value="<%=forum("baslik")%>" maxlength="100"><P>

<!--#INCLUDE file="editor.asp"-->

</td></tr>

<tr><td align="center" valign="center" width="100%">
<textarea onkeyup=textKey(this.form) name="yorum"  cols="150" rows="20">
<%=terseditor(forum("aciklama"))%></textarea>
<P>
<input type="hidden" name="tarih" size="30"   value="<%=(Date)%>">
<input type="submit" name="Submit" value=" Kaydet ">
<input type="reset" name="" value=" Temizle ">
</form>

</td></tr></table>


<%

End If
If gorev="kaydet" Then

if request.form("baslik")="" or request.form("yorum")=""  then
Response.Write "<CENTER><IMG SRC=""forumimg/alert.gif"" WIDTH=""32""  BORDER=""0""><P>"
Response.Write "<B>Lütfen Gerekli Alanlarý Doldurunuz</B><P>"
Response.Write "<a href=""javascript:history.back(1)""><B>Geri gidip tekrar deneyiniz</B></a></CENTER>"
Response.End
End If

If session("yazdim") = True Then
Response.Redirect "default.asp"
Else
session("yazdim")=true
End If


If yer="soru" then
sor = "select  * from  sorular where id="&id&"  "
else
sor = "select  * from cevaplar where id="&id&"  "
End If
forum.Open sor,forumbag,1,3
session("gkodu2")=""

forum("tarih") = temizle(Request.Form ("tarih"))
forum("baslik") = temizle(Request.Form ("baslik"))

aciklama=temizle(Request.Form ("yorum"))
aciklama=editor(aciklama)
forum("aciklama") =aciklama
forum("ipno") = Request.ServerVariables("REMOTE_ADDR")

minsayi = 10000 
maxsayi = 99999
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
guvenlik = Int(sngRandomValue)

forum("guvenlik") =guvenlik

forum.Update


'ADMÝNE BÝLDÝR
email          =emailadresim1
emesaj       ="<B>" &forum("baslik") & "</B><BR>" & aciklama

If yer="soru" then
konu          ="Forumda konu güncellendi"
sil="<A HREF="&websayfam&"/islem.asp?isl=soru&grv=sil&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Sil</b></a> <P>"
onay="<A HREF="&websayfam&"/islem.asp?isl=soru&grv=onay&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Onayla</b></a><P> "

Else
konu          ="Forumda mesaj güncellendi"
sil="<A HREF="&websayfam&"/islem.asp?isl=cevap&grv=sil&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Sil</b></a> <P>"
onay="<A HREF="&websayfam&"/islem.asp?isl=cevap&grv=onay&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Onayla</b></a><P> "
End If

siteyegit="<A HREF="&websayfam&">Siteye Git</a><P>"

emesaj= emesaj & sil & onay & siteyegit

call emailgonder(email,konu,emesaj)


If yer="soru" then
Response.Redirect "default.asp?part=oku&id="&forum("grp")&"&pid="&forum("altgrp")&"&pid1="&forum("sub")&"&urun="&forum("id")&""
Else
Response.Redirect "default.asp?part=oku&id="&forum("grp")&"&pid="&forum("altgrp")&"&urun="&forum("soruid")&""
End If

forum.close
End If



set forum =Nothing
set forum1 =Nothing
set forum2 =Nothing
set efkan =Nothing
set efkan1 =Nothing
%>











