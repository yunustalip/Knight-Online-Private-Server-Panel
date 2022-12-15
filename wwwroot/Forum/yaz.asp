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
id       =kontrol(Request.QueryString("id"))
pid     =kontrol(Request.QueryString("pid"))
pid1     =kontrol(Request.QueryString("pid1"))
urun   =kontrol(Request.QueryString("urun"))
gorev  =temizle(Request.QueryString("gorev"))

If gorev="soru" Or gorev="" then
session("yazdim")=false
%>

<form name="formcevap" onsubmit="return formCheck(this);"  action="yaz.asp?gorev=skaydet" method="post" >
<table width="100%" border="1" bgcolor="" bordercolor="#FFFFFF" align="center" cellpadding="0" cellspacing="0">
<tr height=""><td width="100%" align="center" valign="center">

<%sor = "select  * from  grup where id="&pid&" order by grp asc "
forum.Open sor,forumbag,1,3
grupadi=forum("grp")
forum.close
%>
<BR>
<B><%=grupadi%> Alt  Kategorisinde Yeni Konu Açýyorsunuz</B>
<P>

<INPUT TYPE="hidden" NAME="grp" size=""  value="<%=id%>">
<INPUT TYPE="hidden" NAME="altgrp" size="" value="<%=pid%>">
<INPUT TYPE="hidden" NAME="pid1" size="" value="<%=pid1%>">
<!--#INCLUDE file="editor.asp"-->
</td></tr>

<tr><td align="center" valign="center" width="100%"><BR>
<B>Konu Baþlýðý :&nbsp;</B><INPUT TYPE="text" NAME="Baslik" size="100" maxlength="100"><P>

<B>Sorunuz</B><BR>
<textarea onkeyup=textKey(this.form) name="yorum"  cols="150" rows="20" ></textarea>
<P>
<input type="hidden" name="tarih" size="30"   value="<%=(Date)%>">
<input type="submit" name="Submit" value=" Kaydet ">
<input type="reset" name="" value=" Temizle ">
</form>

</td></tr></table>

<%
End If

If gorev="skaydet" Then
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

sor  = "select * from sorular "
forum.Open sor,forumbag,1,3
forum.addnew
forum("onay") =hemenyayinla
forum("grp") = temizle(Request.Form ("grp"))
forum("altgrp") = temizle(Request.Form ("altgrp"))

If Request.Form ("pid1")<>"" then
forum("sub") = temizle(Request.Form ("pid1"))
End If

forum("tarih") = temizle(Request.Form ("tarih"))
forum("baslik") = temizle(Request.Form ("baslik"))

aciklama=temizle(Request.Form ("yorum"))
aciklama=editor(aciklama)

forum("aciklama") =aciklama
forum("ipno") = Request.ServerVariables("REMOTE_ADDR")
forum("kadi") = session("kadi")
forum("uyeid") = session("uyeid")

minsayi = 10000 
maxsayi = 99999
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
guvenlik = Int(sngRandomValue)

forum("guvenlik") = guvenlik
forum.Update

email          =emailadresim1
konu          ="Foruma soru eklendi "
emesaj       ="<B>" &forum("baslik") & "</B><BR>" & aciklama & "<P>"

sil="<A HREF="&websayfam&"/islem.asp?isl=soru&grv=sil&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Sil</b></a> <P>"
onay="<A HREF="&websayfam&"/islem.asp?isl=soru&grv=onay&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Onayla</b></a><P> "

siteyegit="<A HREF="&websayfam&">Siteye Git</a><P>"

emesaj= emesaj & sil & onay & siteyegit

call emailgonder(email,konu,emesaj)

If hemenyayinla=0 then
Response.Write "<BR><BR><BR><b>Kaydýnýz yönetici onayýndan sonra yayýnlanacaktýr </b><P>"
Response.Write "<meta http-equiv='Refresh' content='2; URL=default.asp'>"
Else

If Request.Form ("pid1")<>"" then
Response.Redirect "default.asp?part=oku&id="&forum("grp")&"&pid="&forum("altgrp")&"&pid1="&forum("sub")&"&urun="&forum("id")&""
else
Response.Redirect "default.asp?part=oku&id="&forum("grp")&"&pid="&forum("altgrp")&"&urun="&forum("id")&""
End If

End If
forum.close

End If


If gorev="cevap"  Then
yer=Request.QueryString("yer")  'ALINITNIN SORUDAN MI CEVAPDAN MI OLDUÐU
alid=Request.QueryString("alid")  'CEVAPTAN ALINTIYSA CEVAP ÝDÝ
session("yazdim")=false
%>

<form name="formcevap" onsubmit="return formCheck(this);" action="yaz.asp?gorev=ckaydet" method="post" >
<table width="800" border="1" bgcolor="" bordercolor="#FFFFFF" align="center" cellpadding="0" cellspacing="0">
<tr height=""><td width="100%" align="center" valign="center">
<%sor = "select  * from  sorular where id="&urun&" order by grp asc "
forum.Open sor,forumbag,1,3
soruadi=forum("baslik")
forum.close
%>
<BR><B><%=soruadi%> Konusuna Cevap Yazýyorsunuz</B>
<P>

<INPUT TYPE="hidden" NAME="grp" size=""  value="<%=id%>">
<INPUT TYPE="hidden" NAME="altgrp" size="" value="<%=pid%>">
<INPUT TYPE="hidden" NAME="pid1" size="" value="<%=pid1%>">
<INPUT TYPE="hidden" NAME="soruid" size="" value="<%=urun%>">

<!--#INCLUDE file="editor.asp"-->
</td></tr>

<tr><td align="center" valign="center" width="100%">

<% if yer="soru" Then 
sor = "select  * from  sorular where id="&urun&" "
forum.Open sor,forumbag,1,3
%>
<B>Alýntý </B><I>(300 karekter alýndý)</I><BR>
<textarea name="alinti" cols="150" rows="2" readonly >

<fieldset style="background:#e1e4f2; BORDER: 0px solid;">
<FONT COLOR="blue"><B>Alýntý :</B>&nbsp;
id:<B><%=urun%></B>&nbsp;
Yazan:<B><%=forum("kadi")%></B>&nbsp;
Tarih:<B><%=forum("tarih")%></B></FONT><BR>
<I><B><%=kucukharf(forum("baslik"))%></B></I>:
<I><%=htmltemizle(Left(forum("aciklama"),300))%>...</I>
</fieldset>

</textarea>

<!-- ALINTI ÖN ÝZLEME -->
<div align="left">
<P>Alýntý Ön izleme
<fieldset style="background:#e1e4f2; BORDER: 0px solid;">
<FONT COLOR="blue"><B>Alýntý :</B>&nbsp;
id:<B><%=urun%></B>&nbsp;
Yazan:<B><%=forum("kadi")%></B>&nbsp;
Tarih:<B><%=forum("tarih")%></B></FONT><BR>
<I><B><%=kucukharf(forum("baslik"))%></B></I>:
<I><%=htmltemizle(Left(forum("aciklama"),300))%>...</I>
</fieldset>
</div>

<% 
forum.close
ElseIf yer="cevap" Then 
sor = "select  * from  cevaplar where id="&alid&"  "
forum.Open sor,forumbag,1,3
%>


<B>Alýntý </B><I>(300 karekter alýndý)</I><BR>
<textarea  name="alinti" cols="150" rows="2" readonly >

<fieldset style="background:#e1e4f2; BORDER: 0px solid;">
<FONT COLOR="blue"><B>Alýntý :</B>&nbsp;
id:<B><%=alid%></B>&nbsp;
Yazan:<B><%=forum("kadi")%></B>&nbsp;
Tarih:<B><%=forum("tarih")%></B></FONT><BR>
<I><B><%=kucukharf(forum("baslik"))%></B></I>:
<I><%=htmltemizle(Left(forum("aciklama"),300))%>...</I>
</fieldset>

</textarea>

<!-- ALINTI ÖN ÝZLEME -->
<P>Alýntý Ön izleme
<div align="left">
<fieldset style="background:#e1e4f2; BORDER: 0px solid;">
<FONT COLOR="blue"><B>Alýntý :</B>&nbsp;
id:<B><%=alid%></B>&nbsp;
Yazan:<B><%=forum("kadi")%></B>&nbsp;
Tarih:<B><%=forum("tarih")%></B></FONT><BR>
<I><B><%=kucukharf(forum("baslik"))%></B></I>:
<I><%=htmltemizle(Left(forum("aciklama"),300))%>...</I>
</fieldset>
</div>

<table>



<%
forum.close
else
End If
%>

<BR><BR>
<B>Cevap Baþlýðý :&nbsp;</B><INPUT TYPE="text" NAME="Baslik" size="100" maxlength="100"><P>


<B>Mesajýnýz</B><BR>
<textarea onkeyup=textKey(this.form) name="yorum"  cols="150" rows="20" ></textarea>
<P>
<input type="hidden" name="tarih" size="30"   value="<%=(Date)%>">
<input type="submit" name="Submit" value=" Kaydet ">
<input type="reset" name="" value=" Temizle ">
</form>
</td></tr></table>
<%
End If

If gorev="ckaydet" Then
Response.Buffer = True 

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

sor  = "select * from cevaplar "
forum.Open sor,forumbag,1,3
session("gkodu2")=""
forum.addnew
forum("onay") =hemenyayinla
forum("grp") = temizle(Request.Form ("grp"))
forum("altgrp") = temizle(Request.Form ("altgrp"))
forum("soruid") = temizle(Request.Form ("soruid"))
forum("tarih") = temizle(Request.Form ("tarih"))
forum("baslik") = temizle(Request.Form ("baslik"))

aciklama=temizle(Request.Form ("yorum"))
aciklama=editor(aciklama)
forum("aciklama") = aciklama
forum("alinti") = Request.Form ("alinti")
'forum("aciklama") =Request.Form ("alinti") & "<p>" & aciklama

forum("ipno") = Request.ServerVariables("REMOTE_ADDR")
forum("kadi") = session("kadi")
forum("uyeid") = session("uyeid")


minsayi = 10000 
maxsayi = 99999
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
guvenlik = Int(sngRandomValue)

forum("guvenlik") = guvenlik

forum.Update

email          =emailadresim1
konu          ="Foruma cevap eklendi"
emesaj       ="<B>" &forum("baslik") & "</B><BR>" & aciklama &"<P>"

sil="<A HREF="&websayfam&"/islem.asp?isl=cevap&grv=sil&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Sil</b></a> <P>"
onay="<A HREF="&websayfam&"/islem.asp?isl=cevap&grv=onay&id="&forum("id")&"&guvenlik="&guvenlik&"><b>Onayla</b></a><P> "

siteyegit="<A HREF="&websayfam&"><B>Siteye Git</B></a><P>"

emesaj= emesaj & sil & onay & siteyegit

call emailgonder(email,konu,emesaj)
forum.close


soruid=temizle(Request.Form ("soruid"))
sor = "select  * from  sorular where id="&soruid&"  "
forum.Open sor,forumbag,1,3
soruadi =forum("baslik")
uyeid   =forum("uyeid")
id =forum("grp")
pid =forum("altgrp")
urun =forum("id")
forum.close


'SORU SAHÝBÝNE MESAJ GÖNDER
sor="SELECT * FROM uyeler WHERE id ="&uyeid&"  "
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
efkan1.close
Else
email          =efkan1("email")
konu          =soruadi & "Sorunuza  cevap eklendi"
emesaj       ="Sayýn " &efkan1("adi") & "<BR><B>" & soruadi & "</B>&nbsp;sorusuna cevap eklendi<P>"
emesaj       =emesaj & "<A HREF='"&websayfam&"'><B>&nbsp;Siteye Git</B></A>&nbsp;"
call emailgonder(email,konu,emesaj)
efkan1.close
End If



'TAKÝP LÝSTESÝNE MESAJ GONDER
sor="SELECT * FROM takip where urun="&soruid&"  "
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
efkan1.close
Else
do while not efkan1.eof 
sor="SELECT * FROM uyeler WHERE id =" &efkan1("uyeid") &"  "
efkan.Open sor,Sur,1,3
    If efkan.eof Then
	Else 'ÜYE YOKSA SÝLÝNDÝYSE
    email          =efkan("email")
    konu          =soruadi & "Sorusuna  cevap eklendi"
    emesaj       ="Sayýn " &efkan("adi") & "<BR><B>" & soruadi & "</B>&nbsp;sorusuna cevap eklendi<P>"
    emesaj       =emesaj & "<A HREF='"&websayfam&"'><B>&nbsp;Siteye Git</B></A>&nbsp;"
     call emailgonder(email,konu,emesaj)
    
    End If
     efkan.close
efkan1.movenext 
loop 
efkan1.close
End If

If hemenyayinla=0 then
Response.Write "<BR><BR><BR><b>Kaydýnýz yönetici onayýndan sonra yayýnlanacaktýr </b><P>"
Response.Write "<meta http-equiv='Refresh' content='2; URL=default.asp'>"
Else

If Request.Form ("pid1")<>"" Then
pid1=Request.Form ("pid1")
Response.Redirect "default.asp?part=oku&id="&id&"&pid="&pid&"&pid1="&pid1&"&urun="&urun&""
Else
Response.Redirect "default.asp?part=oku&id="&id&"&pid="&pid&"&urun="&urun&""
End If

End If
End If




set forum =Nothing
set forum1 =Nothing
set forum2 =Nothing
set efkan =Nothing
set efkan1 =Nothing
%>











