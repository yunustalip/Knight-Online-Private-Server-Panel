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
<BR><BR>
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





<% Response.Buffer = True

'Response.Expires=0
Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage



ip=Request.ServerVariables("REMOTE_ADDR")
sor = "SELECT * FROM yasakli where ip='"& ip &"'"
efkan.open sor, sur, 1, 3
If efkan.eof Then
Else
Response.Write "<script language='JavaScript'>alert('Bu siteye üye giriþi yapmanýz yasaklanmýþtýr');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>" 
End If
efkan.close


gel = Request.ServerVariables("HTTP_REFERER")

gorev=request.querystring("gorev") 



'AKTÝFLEÞTÝRME EMAÝLÝM GELMEDÝ
if gorev="emailgelmedi" then 
If Request.ServerVariables("CONTENT_LENGTH")=0 Then %>
<B>ÜYELÝK AKTÝFLEÞTÝRME EMAÝLÝM GELMEDÝ</B>
<form action="default.asp?part=uyegorev&gorev=emailgelmedi" method="POST" >
<table width="400" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="5">
<tr><td align="" width="35%"><B>Email Adresiniz</B></td>
<td align="" width="65%">*<input name="email" size="30" maxlength="50"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="Tamam">
</td></tr></table></form>
<% 
Else
If Request.Form ("email")="" Then
hataver("Lütfen üye olurken kullandýðýnýz email adresini yazýnýz.")
Else
formemail =temizle(Request.Form ("email"))
sor="select * from uyeler where email='"&formemail&"'  "
efkan.Open sor,Sur,1,3
if efkan.eof or efkan.bof Then
hataver("Bu emaile ait bir kayýt bulunamadý")
ElseIf  efkan("onay")=1 Then
bilgiver("Üyeliðiniz zaten aktif")
ElseIf efkan("onay")<>1 Then
emesaj = "Sayýn " &efkan("adi")& " " & Now() & "<BR>"
emesaj =emesaj & websayfam & " Sitemize yaptýðýnýz üyelik baþvurusunun  aktif olabilmesi için verilen  linke týklayýnýz.<BR><B>4 gün</B> sonunda üyeliðinizi aktifleþtirmezseniz.Üyeliðiniz silinecektir.<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")&"'>Üyeliðimi Aktif Et</A><P> "
emesaj = emesaj & " Eðer verilen linkten dönemiyorsanýz aþaðýdaki linki tarayýcýnýzýn adres satýrýna yapýþtýrarak iþleminizi tamamlayabilirsiniz<P> "
emesaj = emesaj & websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")
email        = efkan("email")
konu        =" Üyelik Aktivasyonu "
emesaj     =emesaj
call emailgonder(email,konu,emesaj)
bilgiver("Email Adresinize Aktivasyon linki gönderildi.<BR>Lütfen emaillerinizi kontrol ediniz.<BR>4 gün içinde aktifleþtirilmeyen üyelikler silinecektir.")
End If
efkan.close
End If
End If
End If






'GÝRÝÞ FORM 
if gorev="girisform" then 
gkod1  =kodver2(gkod) 
%>
<form action="uyegorev.asp?gorev=kontrol" method="POST">
<table  width="300" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="5">
<tr><td colspan="2" align="center"><B>Sadece Üyeler</B></td><td></tr>
<tr><td width="40%">Kullanýcý Adý</td><td width="60%">
<input name="kadi" size="20" maxlength="10"></td></tr>
<tr><td>Parola</td><td><input type="password" name="sifre" size="20" maxlength="10"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="Üye Giriþi"></td></tr>
<tr><td colspan="2" align="center">
<A HREF="?part=uyegorev&gorev=uyeol">Üye Olacaðým</A> |
<A HREF="?part=uyegorev&gorev=unuttum">Þifremi Unuttum</A></td></tr>
</table></form>

<% 

End If 

'KONTROL
if gorev="kontrol" then 
formkadi = temizle(Request.Form("kadi"))
formsifre = encode(temizle(Request.Form("sifre")))
sor="select * from uyeler where kadi='"&formkadi&"' and sifre='"&formsifre&"' and onay=1 "
efkan.Open sor,Sur,1,3
if efkan.eof or efkan.bof then
Response.Write "<script language='JavaScript'>alert('Kullanýcý adý veya þifre yanlýþ lütfen tekrar deneyin...');</script>"
Response.Write "<meta http-equiv='Refresh' content='1; URL=default.asp?part=uyegorev&gorev=girisform'>"
Response.End
efkan.close
End If

Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")

efkan("hit")=efkan("hit") + 1
efkan("ipno") = Request.ServerVariables("REMOTE_ADDR") 
efkan.update
id=efkan("id")
efkan.close

'YÖNETÝCÝ ÝSE YÖNETÝCÝ SESSÝON AKTÝF
sor="select * from uyeler where id = "& Session("uyeid") &"   "
efkan.Open sor,Sur,1,3
if efkan("admin") = 1  then
Session("efkanlogin") = True
End If
efkan.close

'OKUNMAYAN MESAJ VAR MI EN SON GÝRÝÞ TEN SONRA
sor="select * from mesaj where kime="&Session("uyeid")&" and okundu = 0 and kimesildi <>1 "
efkan.Open sor,Sur,1,3
mesajvar = efkan.recordcount
if mesajvar >0  then
Response.Write "<script language='JavaScript'>alert('Okunmayan mesaj/mesajlar var..');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyemesaj&gorev=gelen'>" 
End If
efkan.close

'SON TARÝHÝ GÜNCELLE ONLÝNE OLAYI ÝÇÝN
Session.LCID = 1055
DefaultLCID = Session.LCID 

sontarih=Now()
sor="select * from uyeler where id="&Session("uyeid")&" "
efkan.Open sor,Sur,1,3
efkan("sontarih")=sontarih
efkan.Update
efkan.close
'Response.Write "<BR><BR><BR><BR><b>Hoþgeldiniz&nbsp;"&formkadi&"....</b><br>"
'Response.Write "<b>Þimdi Ana Sayfaya yönlendiriliyorsunuz</b><br>"
'Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>" 
Response.Redirect "default.asp"
End If





'ÜYE OLMA FORMU
if gorev="uyeol" then 
gkod1  =kodver2(gkod) %>
<B>ÜYELÝK FORMUNU LÜTFEN DOLDURUNUZ...</B>

<form name="formcevap" onsubmit="return formCheck(this);" action="default.asp?part=uyegorev&gorev=uyeoltamam" method="POST" >

<iframe src="kurallar.asp" width="500" height="100" marginwidth="0" marginheight="0" hspace="0" vspace="0" frameborder="0" scrolling="yes">
 </iframe>

<table background="" width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="4">

<tr><td width="30%">Avatar Seçiniz</td><td width="70%">
<%sor="SELECT * FROM avatar  "
efkan.Open sor,Sur,1,3%>
<select name="avatar" size="7" class="form" onChange="(resim.src = avatar.options[avatar.selectedIndex].value)">
<option selected="avatar/01.jpg">01.jpg</option>
<% do while not efkan.eof  %>
<option value="avatar/<%=efkan("avatar")%>"><%=efkan("avatar")%></option>
<% efkan.movenext 
loop %>
</select>
<img src="avatar/01.jpg" width="150" height="110" name="resim">
<%efkan.close%>
<BR>
<!-- FOTO YUKLEME SERBESTSE -->
<% if fotoyukleme=1 Then %>Üye olduktan sonra kendi resminizi yükleyebilirsiniz...<% End If %>
</td></tr>

<tr><td>Güvenlik Kodu</td><td>
<B><FONT COLOR="red"><%=gkod1%></FONT></B>
*<input type="text" name="gkod1" size="20" maxlength="20" onkeypress="return SayiKontrol(this);">
</td></tr>

<tr><td>Kullanýcý Adý</td><td>*<input name="kadi" size="20" maxlength="20"></td></tr>
<tr><td>Þifreniz</td><td>*<input type="password" name="sifre" size="20" maxlength="20"></td></tr>
<tr><td>Adýnýz</td><td>*<input type="text" name="adi" size="50" maxlength="50"></td></tr>
<tr><td>Email Adresiniz</td><td>*<input type="text" name="email" size="50" maxlength="50"></td></tr>
<tr><td>Web Sayfanýz</td><td><input type="text" name="url" size="50" maxlength="50" value="http://"></td></tr>


<tr><td>Msn Messenger</td><td><input type="text" name="msn" size="50" maxlength="50"></td></tr>
<tr><td>Yahoo</td><td><input type="text" name="yahoo" size="50" maxlength="50"></td></tr>
<tr><td>Icq</td><td>
<input type="text" name="icq" size="50" maxlength="50" onkeypress="return SayiKontrol(this);"></td></tr>

<tr><td>Mesleðiniz</td><td>
<SELECT  name="meslek"> 
<OPTION selected value="">Mesleðinizi Seçin</OPTION> 
<OPTION value="Çalýþmýyorum">Çalýþmýyorum</OPTION>
<OPTION value="Akademisyen, Öðretmen">Akademisyen, Öðretmen</OPTION> 
<OPTION value="Avukat">Avukat</OPTION>
<OPTION value="Bankacý">Bankacý</OPTION> 
<OPTION value="Bilgisayar, Internet">Bilgisayar, Internet</OPTION>
<OPTION value="Danýþman">Danýþman</OPTION>
<OPTION value="Doktor ">Doktor</OPTION> 
<OPTION value="Emekli ">Emekli</OPTION>
<OPTION value="Ev Hanýmý ">Ev Hanýmý</OPTION>
<OPTION value="Finasman, Muhasebe ">Finasman, Muhasebe</OPTION> 
<OPTION value="Fotoðrafçý ">Fotoðrafçý</OPTION>
<OPTION value="Gazeteci ">Gazeteci</OPTION>
<OPTION value="Grafiker ">Grafiker</OPTION>
<OPTION value="Manken,Fotomodel ">Manken,Fotomodel</OPTION>
<OPTION value="Memur ">Memur</OPTION> 
<OPTION value="Mühendis ">Mühendis</OPTION>
<OPTION value="Öðrenci ">Öðrenci</OPTION> 
<OPTION value="Politikacý ">Politikacý</OPTION>
<OPTION value="Psikolog ">Psikolog</OPTION>
<OPTION value="Reklamcý ">Reklamcý</OPTION>
<OPTION value="Sanatçý ">Sanatçý</OPTION> 
<OPTION value="Satýþ, Pazarlama ">Satýþ, Pazarlama</OPTION>
<OPTION value="Serbest Meslek, Ýþ Sahibi ">Serbest Meslek, Ýþ Sahibi</OPTION>
<OPTION value="Sporcu ">Sporcu</OPTION>
<OPTION value="Teknik Eleman ">Teknik Eleman</OPTION>
<OPTION value="Üst Düzey Yönetici ">Üst Düzey Yönetici</OPTION> 
<OPTION value="Diðer ">Diðer</OPTION>
 </SELECT> 
</td></tr>


<tr><td>Yaþýnýz</td><td>
<select name="yas">
<option selected>16</option>
<% i=16
do while i<60
i=i+1 %>
<option><%=i%></option>
<% loop %>
</select> 
</td></tr>

<tr><td>Þehir</td><td>
<SELECT name="sehir" class="Input" >
          <OPTION value="ADANA">ADANA</OPTION>
          <OPTION value="ADIYAMAN">ADIYAMAN</OPTION>
          <OPTION value="AFYON">AFYON</OPTION>
          <OPTION value="AÐRI">AÐRI</OPTION>
          <OPTION value="AKSARAY">AKSARAY</OPTION>
          <OPTION value="AMASYA">AMASYA</OPTION>
          <OPTION value="ANKARA">ANKARA</OPTION>
          <OPTION value="ANTALYA">ANTALYA</OPTION>
          <OPTION value="ARDAHAN">ARDAHAN</OPTION>
          <OPTION value="ARTVÝN">ARTVÝN</OPTION>
          <OPTION value="AYDIN">AYDIN</OPTION>
          <OPTION value="BALIKESÝR">BALIKESÝR</OPTION>
          <OPTION value="BARTIN">BARTIN</OPTION>
          <OPTION value="BATMAN">BATMAN</OPTION>
          <OPTION value="BAYBURT">BAYBURT</OPTION>
          <OPTION value="BÝLECÝK">BÝLECÝK</OPTION>
          <OPTION value="BÝNGÖL">BÝNGÖL</OPTION>
          <OPTION value="BÝTLÝS">BÝTLÝS</OPTION>
          <OPTION value="BOLU">BOLU</OPTION>
          <OPTION value="BURDUR">BURDUR</OPTION>
          <OPTION value="BURSA">BURSA</OPTION>
          <OPTION value="ÇANAKKALE">ÇANAKKALE</OPTION>
          <OPTION value="ÇANKIRI">ÇANKIRI</OPTION>
          <OPTION value="ÇORUM">ÇORUM</OPTION>
          <OPTION value="DENÝZLÝ">DENÝZLÝ</OPTION>
          <OPTION value="DÝYARBAKIR">DÝYARBAKIR</OPTION>
          <OPTION value="DÜZCE">DÜZCE</OPTION>
          <OPTION value="EDÝRNE">EDÝRNE</OPTION>
          <OPTION value="ELAZIÐ">ELAZIÐ</OPTION>
          <OPTION value="ERZÝNCAN">ERZÝNCAN</OPTION>
          <OPTION value="ERZURUM">ERZURUM</OPTION>
          <OPTION value="ESKÝÞEHÝR">ESKÝÞEHÝR</OPTION>
          <OPTION value="GAZÝANTEP">GAZÝANTEP</OPTION>
          <OPTION value="GÝRESUN">GÝRESUN</OPTION>
          <OPTION value="GÜMÜÞHANE">GÜMÜÞHANE</OPTION>
          <OPTION value="HAKKARÝ">HAKKARÝ</OPTION>
          <OPTION value="HATAY" >HATAY</OPTION>
          <OPTION value="IÐDIR">IÐDIR</OPTION>
          <OPTION value="ISPARTA">ISPARTA</OPTION>
          <OPTION value="ÝÇEL">ÝÇEL</OPTION>
          <OPTION value="ÝSTANBUL" selected>ÝSTANBUL</OPTION>
          <OPTION value="ÝZMÝR">ÝZMÝR</OPTION>
          <OPTION value="KAHRAMANMARAÞ">KAHRAMANMARAÞ</OPTION>
          <OPTION value="KARABÜK">KARABÜK</OPTION>
          <OPTION value="KARAMAN">KARAMAN</OPTION>
          <OPTION value="KARS">KARS</OPTION>
          <OPTION value="KASTAMONU">KASTAMONU</OPTION>
          <OPTION value="KAYSERÝ">KAYSERÝ</OPTION>
          <OPTION value="KIBRIS">KIBRIS</OPTION>
          <OPTION value="KIRIKKALE">KIRIKKALE</OPTION>
          <OPTION value="KIRKLARELÝ">KIRKLARELÝ</OPTION>
          <OPTION value="KIRÞEHÝR">KIRÞEHÝR</OPTION>
          <OPTION value="KÝLÝS">KÝLÝS</OPTION>
          <OPTION value="KOCAELÝ">KOCAELÝ</OPTION>
          <OPTION value="KONYA">KONYA</OPTION>
          <OPTION value="KÜTAHYA">KÜTAHYA</OPTION>
          <OPTION value="MALATYA">MALATYA</OPTION>
          <OPTION value="MANÝSA">MANÝSA</OPTION>
          <OPTION value="MARDÝN">MARDÝN</OPTION>
          <OPTION value="MUÐLA">MUÐLA</OPTION>
          <OPTION value="MUÞ">MUÞ</OPTION>
          <OPTION value="NEVÞEHÝR">NEVÞEHÝR</OPTION>
          <OPTION value="NÝÐDE">NÝÐDE</OPTION>
          <OPTION value="ORDU">ORDU</OPTION>
          <OPTION value="OSMANÝYE">OSMANÝYE</OPTION>
          <OPTION value="RÝZE">RÝZE</OPTION>
          <OPTION value="SAKARYA">SAKARYA</OPTION>
          <OPTION value="SAMSUN">SAMSUN</OPTION>
          <OPTION value="SÝÝRT">SÝÝRT</OPTION>
          <OPTION value="SÝNOP">SÝNOP</OPTION>
          <OPTION value="SÝVAS">SÝVAS</OPTION>
          <OPTION value="ÞANLIURFA">ÞANLIURFA</OPTION>
          <OPTION value="ÞIRNAK">ÞIRNAK</OPTION>
          <OPTION value="TEKÝRDAÐ">TEKÝRDAÐ</OPTION>
          <OPTION value="TOKAT">TOKAT</OPTION>
          <OPTION value="TRABZON">TRABZON</OPTION>
          <OPTION value="TUNCELÝ">TUNCELÝ</OPTION>
          <OPTION value="UÞAK">UÞAK</OPTION>
          <OPTION value="VAN">VAN</OPTION>
          <OPTION value="YALOVA">YALOVA</OPTION>
          <OPTION value="YOZGAT">YOZGAT</OPTION>
          <OPTION value="ZONGULDAK">ZONGULDAK</OPTION>
        </SELECT> 
</td></tr>

<tr><td colspan="2" align="center">
Ýmzanýz <I>(500 Karekter)</I><P>
<!--#INCLUDE file="editor.asp"-->
<P>
<TEXTAREA  onkeyup=textKey(this.form) name="yorum" ROWS="8" COLS="80"></TEXTAREA>
</td></tr>

<tr><td colspan="2" align="center">
<input type="hidden" name="tarih" size="30"   value="<%=(Date)%>">
<input type="submit" value="Üye Ol">&nbsp;&nbsp;
<input type="reset" value="Temizle">
</td></tr></table></form>
<%End If %>

<!-- ÜYE OLMA FORUMU ÝÞLENÝYOR -->
<% 
if gorev="uyeoltamam" then 

'GÜVENLÝK KODU KONTROL
if  temizle(Request.Form("gkod1")) <> trim(session("gkodu2")) Then
Response.Write  "<BR><BR><BR><center>Güvenlik kodu yazýlmamýþ veya yanlýþ <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

if request.form("kadi")="" or request.form("sifre")="" or request.form("email")="" or request.form("adi")="" then
Response.Write "<BR><BR><BR><center>Lütfen iþaretli alanlarý doldurunuz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

if emailkontrol(Request.Form ("email"))=false then
Response.Write "<BR><BR><BR><center>Geçerli Email adresi kullanýn <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

formkadi = temizle(Request.Form("kadi"))
kadisay=Len(formkadi) 
formsifre = temizle(Request.Form("sifre"))
sifresay=Len(formsifre) 
email=temizle(Request.Form("email"))

if kadisay < 4 or sifresay <4 then 
Response.Write "<BR><BR><BR><center>En az <B>4</B> karekter uzunluðunda kullanýcý adý ve þifre kullanýnýz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

'BU KADÝ AYNI OLAN VARMI
sor="select * from uyeler where kadi='"&formkadi&"' OR email='"&email&"'   "
efkan.Open sor,sur,1,3
adet=efkan.recordcount
if adet > 0 Then
Response.Write "<BR><BR><BR><center>Bu kullanýcý adý veya email kullanýlýyor baþka bir kullanýcý adý veya email adresi deneyiniz... <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
else
efkan.close

sor  = "select * from uyeler "
efkan.Open sor,sur,1,3
session("gkodu2")=""
efkan.AddNew

Session.LCID = 1055
DefaultLCID = Session.LCID 
sontarih=Now()
efkan("sontarih")=sontarih
efkan("kadi") = Temizle(Request.Form ("kadi"))
efkan("sifre") = encode(Temizle(Request.Form ("sifre")))
efkan("adi") = Temizle(Request.Form ("adi"))
efkan("email") = Temizle(Request.Form ("email"))
efkan("url") = Temizle(Request.Form ("url"))
efkan("meslek") = Temizle(Request.Form ("meslek"))
efkan("yas") = Temizle(Request.Form ("yas"))
efkan("sehir") = Temizle(Request.Form ("sehir"))
'efkan("soru") = Temizle(Request.Form ("soru"))
'efkan("cevap") = Temizle(Request.Form ("cevap"))
efkan("msn") = Temizle(Request.Form ("msn"))
efkan("yahoo") = Temizle(Request.Form ("yahoo"))
efkan("icq") = Temizle(Request.Form ("icq"))
efkan("tarih") = Temizle(Request.Form ("tarih"))

avatar =Temizle(Request.Form ("avatar"))
avatar =Replace (avatar ,"avatar/","",1,-1,1)
efkan("avatar") =avatar

imza=suz(Temizle(Request.Form ("yorum")))
imza=left(imza,500)
efkan("imza") =imza

efkan("ipno") = Request.ServerVariables("REMOTE_ADDR") 
efkan.Update


'EMAÝLLE AKTÝVASYON AÇIKSA 
If emaildogrulama=1 Then

minsayi = 10000 'seçilecek sayýnýn alt sýnýrý
maxsayi = 99999 'seçilecek sayýnýn üst sýnýrý
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
aktifkod = Int(sngRandomValue)
efkan("onay") =aktifkod 
efkan.Update
id=efkan("id")

emesaj =websayfam & "&nbsp;Sitemize yaptýðýnýz üyelik baþvurusunun  aktif olabilmesi için verilen  linke týklayýnýz<P><B>4 gün</B> sonunda üyeliðinizi aktifleþtirmezseniz.Üyeliðiniz silinecektir.<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&aktifkod&"&id="&id&"'>Üyeliðimi Aktif Et</A> "


emesaj = emesaj & " <P>Eðer verilen linkten dönemiyorsanýz aþaðýdaki linki tarayýcýnýzýn adres satýrýna yapýþtýrarak iþleminizi tamamlayabilirsiniz<P> "
emesaj = emesaj & websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&aktifkod&"&id="&id


email          =efkan("email")
konu          ="Üyelik Aktivasyonu"
emesaj       =emesaj
call emailgonder(email,konu,emesaj)

Response.Write "<BR><BR><BR><b>Sayýn&nbsp;"&efkan("adi")&"....</b><P>"
Response.Write "<b>Email adresinize aktivasyon linki gönderildi </b><P>"
Response.Write "<b>Lütfen emaillerinizi kontrol ediniz..... </b><P>"
Response.Write "<b>4 gün içinde aktifleþmeyen uyelik kayýtlarý silinecektir.</b><P>"
Response.Write "<b>Þimdi Ana Sayfaya yönlendiriliyorsunuz</b><P>"
Response.Write "<meta http-equiv='Refresh' content='6; URL=default.asp'>"

'ÜYELÝK SERBESTSE
Else
efkan("onay") =1
efkan.Update

'ÜYE LOGÝN OLUYOR
Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")
Response.Write "<BR><BR><BR><b>Hoþgeldiniz&nbsp;"&efkan("adi")&"....</b><br>"
Response.Write "<b>Þimdi Ana Sayfaya yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='4; URL=default.asp'>" 
End If

'EMAÝL LÝSTESÝNE ÝÞLENÝYOR
sor = "Select * from emaillistesi"
efkan1.Open sor,sur,1,3
efkan1.AddNew
efkan1("tarih") = Request.Form ("tarih")
efkan1("adi") =temizle(Request.Form ("adi"))
efkan1("email") =temizle(Request.Form ("email"))
efkan1.Update
efkan1.close

efkan.close
End If 
End If 



'AKTÝVASYON 
if gorev="aktivasyon" then 
kod=kontrol(temizle(request.querystring("kod"))) 
id=kontrol(temizle(request.querystring("id"))) 

sor = "Select * from uyeler where onay = "&kod&" and id="&id&" " 
efkan.Open sor,Sur,1,3
if efkan.eof or efkan.bof Then
Response.Write "<B>Böyle bir Aktivasyon kodu üretilmedi</B><P>"
Response.Write "<b>Þimdi Ana Sayfaya yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='2; URL=default.asp'>" 
Else
efkan("onay") =1
efkan.Update
'ÜYE LOGÝN OLUYOR
Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")
Response.Write "<BR><BR><BR><b>Hoþgeldiniz&nbsp;"&efkan("adi")&"....</b><br>"
Response.Write "<b>Þimdi Ana Sayfaya yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='4; URL=default.asp'>" 
End If
efkan.close
End If



'///////////////// AKTÝVASYONSUZLARA TOPLU AKTÝVASYON EMAÝLÝ  //////////////////////
if gorev="aktivasyonsuzlar" then 
sor = "Select * from uyeler where onay<>1  "
efkan.Open sor,sur,1,3

do while not efkan.eof 
emesaj = "Sayýn " &efkan("adi")& " " & Now() & "<BR>"
emesaj =emesaj & websayfam & " Sitemize yaptýðýnýz üyelik baþvurusunun  aktif olabilmesi için verilen  linke týklayýnýz.<BR><B>4 gün</B> sonunda üyeliðinizi aktifleþtirmezseniz.Üyeliðiniz silinecektir.<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")&"'>Üyeliðimi Aktif Et</A><P> "
emesaj = emesaj & " Eðer verilen linkten dönemiyorsanýz aþaðýdaki linki tarayýcýnýzýn adres satýrýna yapýþtýrarak iþleminizi tamamlayabilirsiniz<P> "
emesaj = emesaj & websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")

email        = efkan("email")
konu        =" Üyelik Aktivasyonu "
emesaj     =emesaj
call emailgonder(email,konu,emesaj)
efkan.movenext 
loop 
efkan.close
Response.Write "Üyeliðini aktif etmeyenlere tekrar email gönderildi"
End If



' ÜYE BÝLGÝ 
if gorev="uyebilgi" then 
Response.Buffer = True 
If Session("uyelogin")=True <> True Then 
Response.Redirect ("default.asp?part=uyegorev&gorev=girisform")
Response.End
End If

id= kontrol(temizle(request.querystring("id")))
sor = "Select * from uyeler where id = "&id&" " 
efkan.Open sor,Sur,1,3
if efkan.eof or efkan.bof then
Response.Write "<BR><BR><B>BU ÜYE SÝLÝNDÝ</B>" 
Response.End
End If
%>
<div align="center">
<table background="" width="80%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="4">
<tr><td colspan="2" align="center" valign="center" width="100%" >
<B>ÜYE BÝLGÝLERÝ</B>
</TD></tr>


<tr><td colspan="2" align="center" valign="center" width="100%" >
<%
sor = "Select * from uyeresim where uyeid = "&id&" " 
efkan1.Open sor,Sur,1,3
if efkan1.eof or efkan1.bof Then%>
<IMG SRC="avatar/<%=efkan("avatar")%>" WIDTH="150"  BORDER="0" ALT="">
<%efkan1.close
Else
Response.Write "<CENTER>Resimleri büyültmek için resime týklayýnýz</CENTER><BR>"
for i=1 to uyeresimadet
if efkan1.eof then exit for%>
<A HREF="uyeler/<%=efkan1("uyeresim")%>" target="_blank">
<IMG SRC="uyeler/<%=efkan1("uyeresim")%>" WIDTH="100" HEIGHT="100" BORDER="0" ALT=""></A>

<!-- ADMÝN LOG ÝSE RESÝM SÝL BUTONU -->
<% If Session("efkanlogin")=True Then %>
<P><A HREF="?part=uyegorev&gorev=resimsil&id=<%=efkan1("id")%>">Resimi sil</a>
<% End If%>

<%
efkan1.movenext
Next
efkan1.close
End If
%>
<P>
<A HREF="?part=uyemesaj&gorev=yaz&id=<%=efkan("id")%>&kime=<%=efkan("kadi")%>">
<IMG SRC="images/yaz.gif" WIDTH="15" HEIGHT="15" BORDER=0 ALT="">
<B>Bu üyeye mesaj gönder</B></A>

<!-- ADMÝN LOG ÝSE UYE SÝL BUTONU -->
<% If Session("efkanlogin")=True Then %>
<P>
<A HREF="?part=uyegorev&gorev=bilgilerim&id=<%=efkan("id")%>">
<IMG SRC="forumimg/degistir.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="Üye Bilgilerini deðiþtir"></a>

<A HREF="?part=uyegorev&gorev=uyesil&id=<%=efkan("id")%>">
<IMG SRC="forumimg/uyesil.gif" WIDTH="46" HEIGHT="17" BORDER="0" ALT="Bu üyeyi sil"></a>

<A HREF="?part=uyegorev&gorev=yasakla&id=<%=efkan("id")%>">
<IMG SRC="forumimg/yasakla.gif" WIDTH="42" HEIGHT="17" BORDER="0" ALT="Bu üyeye siteyi yasakla"></a>
<% End If%>

</TD></tr>

<tr><td width="30%"><B>Ad Soyad</B></td><td width="70%">
<!--EÐER YÖNETÝCÝ ÝSE -->
<% If  efkan("admin") =1  then%> 
<IMG SRC="images/admin.gif" WIDTH="20" HEIGHT="20" BORDER=0 ALT="Yönetici">
<%End If%>
<%=efkan("adi")%>&nbsp;</td></tr>

<tr><td><B>Üye Adi</B></td><td><%=efkan("kadi")%></td></tr>
<tr><td><B>Web Sitesi</B></td>
<td><A HREF="<%=efkan("url")%>" target="_blank"><%=efkan("url")%>&nbsp;</A></td></tr>
<tr><td><B>Yaþ</B></td><td><%=efkan("yas")%>&nbsp;</td></tr>
<tr><td><B>Þehir</B></td><td><%=efkan("sehir")%>&nbsp;</td></tr>
<tr><td><B>Meslek</B></td><td><%=efkan("meslek")%>&nbsp;</td></tr>
<tr><td><B>Msn</B></td><td><%=efkan("msn")%>&nbsp;</td></tr>
<tr><td><B>Yahoo</B></td><td><%=efkan("yahoo")%>&nbsp;</td></tr>
<tr><td><B>Icq</B></td><td><%=efkan("icq")%>&nbsp;</td></tr>
<tr><td><B>Üye Olma Tarihi</B></td><td><%=efkan("tarih")%></td></tr>
<tr><td><B>Son Giriþ</B></td><td><%=efkan("sontarih")%></td></tr>
<tr><td><B>Hit</B></td><td><%=efkan("hit")%></td></tr>
<tr><td><B>Durumu</B></td><td>
<%efkan.close
sor="SELECT * FROM uyeler WHERE id = "&id&"   "
efkan.Open sor,Sur,1,3
Session.LCID = 1055
DefaultLCID = Session.LCID 
zaman=datediff("n",efkan("sontarih"),now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif  WIDTH=11  BORDER=0 ALT=offline>&nbsp;Online Deðil" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>&nbsp;Þu an Online"
End If
efkan.close
%>
</td></tr>


<!-- AÇTIÐI KONULAR -->
<tr><td colspan="2" align="center" valign="center" width="100%" ><B>Açtýðý Son 10 Konu</B></td></tr>

<%'BU UYENÝN KONU SAYISI
sor = "Select * from sorular where onay=1 and uyeid="&id&" order by id desc "  
forum.Open sor,forumbag,1,3
soruadet=forum.recordcount
If soruadet=0 Then
Response.Write "<tr><td colspan=2>Bu üyemiz hiç konu açmadý.</td></tr>"
Else
say = 0
Do While say =< 10 And Not  forum.eof %>
<tr><td colspan="2" align="left" valign="center" width="100%" >
<IMG SRC="images/yaz.gif" WIDTH="15" HEIGHT="15" BORDER="0" ALT="">
<%=forum("tarih")%>
<IMG SRC="images/blank.gif" WIDTH="9" HEIGHT="7" BORDER="0" ALT="">
<A HREF="?part=oku&id=<%=forum("grp")%>&pid=<%=forum("altgrp")%>&urun=<%=forum("id")%>">
<%=forum("baslik")%></a><BR>
</td></tr>
<%
say=say+1
forum.movenext 
loop 
End If
forum.close
%>


<!-- MESAJLARI -->
<tr><td colspan="2" align="center" valign="center" width="100%" ><B>Býraktýðý Son 10 Mesaj</B></td></tr>
<%'BU UYENÝN MESAJ SAYISI
sor = "Select * from cevaplar where onay=1 and uyeid="&id&" order by id desc "  
forum.Open sor,forumbag,1,3
cevapadet=forum.recordcount
If cevapadet=0 Then
Response.Write "<tr><td colspan=2>Bu üyemiz hiç mesaj yazmadý.</td></tr>"
Else
say = 0
Do While say =< 10 And Not  forum.eof %>
<tr><td colspan="2" align="left" valign="center" width="100%" >
<IMG SRC="images/yaz.gif" WIDTH="15" HEIGHT="15" BORDER="0" ALT="">
<%=forum("tarih")%>
<IMG SRC="images/blank.gif" WIDTH="9" HEIGHT="7" BORDER="0" ALT="">
<A HREF="?part=oku&id=<%=forum("grp")%>&pid=<%=forum("altgrp")%>&urun=<%=forum("soruid")%>">
<%=forum("baslik")%></a><BR>
</td></tr>
<%
say=say+1
forum.movenext 
loop 
End If
forum.close
%>

</table>
<% End If %>



<%
'/////////////////////  ÜYELER DÖK /////////////////////////
if gorev="uyeler" Or gorev="" then 
If Session("uyelogin")=True <> True Then 
Response.Redirect ("default.asp?part=uyegorev&gorev=girisform")
Response.End
End If %>

<A HREF="?part=uyegorev&gorev=uyeler"><B>Tüm Üyeleri Göster</B></A>
<table width="95%" bgcolor="" bordercolor="#CCFFFF" border="1" cellspacing="0" cellpadding="5">
<tr bgcolor=""><td align="center">
<form name="uyedok"><select name="menu" onChange="location=document.uyedok.menu.options[document.uyedok.menu.selectedIndex].value;" value="Sýrala">
<option value="?part=uyegorev&gorev=uyeler">Kayýtlarý Sýrala </option>
<option value="?part=uyegorev&gorev=uyeler&diz=id">Yeni Eklenenler</option>
<option value="?part=uyegorev&gorev=uyeler&diz=hit">En Çok Gelenler</option>
<option value="?part=uyegorev&gorev=uyeler&diz=eski">Eskiden Yeniye</option>
<option value="?part=uyegorev&gorev=uyeler&diz=sehir">Þehire Göre</option>
<option value="?part=uyegorev&gorev=uyeler&diz=kadi">Kullanýcý Adýna Göre</option>
<option value="?part=uyegorev&gorev=uyeler&diz=sontarihe">Son Giriþ yapanlar</option>
<% If Session("efkanlogin")=True Then %>
<option value="?part=uyegorev&gorev=uyeler&diz=onaysiz">Onay Bekleyenler</option>
<% End If %></select></td></form>

<!-- ARAMA BÖLÜMÜ -->
<FORM method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?part=uyegorev&gorev=uyeler">
<td align="center">
<input type="text" name="ara" size="20" maxlength="20">&nbsp;&nbsp;<input type="submit" value=" Ara ">
</td></FORM></tr></table>


<%
diz       =trim(Request("diz"))  'DÝZME SEÇENEKLERÝ
ara      =temizle(Request("ara"))



If diz="id" then
sor = "Select * from uyeler  order by id desc"
mesaj="YENÝ ÜYELERÝMÝZ "
ElseIf diz="onaysiz" then
sor = "Select * from uyeler where onay<>1 order by id desc"
mesaj="ONAY BEKLEYENLER "
ElseIf diz="eski" then
sor = "Select * from uyeler  order by id asc"
mesaj="ESKÝ ÜYELERÝMÝZ "
ElseIf diz="sehir" then
sor = "Select * from uyeler  order by sehir asc"
mesaj="ÞEHÝRE GÖRE DÝZÝLDÝ "
ElseIf diz="kadi" then
sor = "Select * from uyeler  order by kadi asc"
mesaj="KULLANICI ADINA GÖRE DÝZÝLDÝ "
ElseIf diz="sontarihe" then
sor = "Select * from uyeler  order by sontarih desc"
mesaj="SON GÝRÝÞE GÖRE DÝZÝLDÝ"
ElseIf diz="hit" then
sor = "Select * from uyeler order by hit desc"
mesaj="EN ÇOK GÝRÝÞ YAPANLAR"
ElseIf ara<>"" then
sor = "select * from uyeler WHERE onay=1 AND kadi like '%"&ara&"%' OR  onay=1 AND email like '%"&ara&"%' order by id desc "
mesaj="ARAMA SONUCU"
Else 
sor = "Select * from uyeler order by id desc"
mesaj="YENÝ ÜYELERÝMÝZ "
End If

efkan.Open sor,Sur,1,3
adet=efkan.recordcount

if efkan.eof or efkan.bof then
hataver("Kayýt Bulunamadý")
else

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>

<P><B><%=mesaj%>&nbsp;</B><BR> Toplam <%=adet%> kayýt
<table background="" width="99%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="3">
<tr bgcolor="<%=bgcolor1%>">
<td align="center" width="5%" ><B>Adý </B></TD>
<td align="center" width="1%" ><B>Online</B></TD>
<td align="center" width="1%" ><B>Konu<BR>mesaj</B></TD>
<td align="center" width="1%" ><B>Hit</B></TD>
<td align="center" width="5%" ><B>Þehir</B></TD>
<td align="center" width="5%" ><B>Son Giriþ</B></TD>

<% If Session("efkanlogin")=True Then %>
<td align="center" width="5%" ><B>Ýþlem</B></TD>
<% End If %>
</tr>

<% renk = 0
efkan.pagesize =50
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit for

if renk mod 2 then
bgcolor = bgcolor1
else
bgcolor = bgcolor2
End If %>
<tr  bgcolor="<%=bgcolor%>"><td align="left">
<!--EÐER YÖNETÝCÝ ÝSE -->
<% If  efkan("admin") =1  then%> 
<IMG SRC="images/admin.gif" WIDTH="20" HEIGHT="20" BORDER=0 ALT="Yönetici">
<%End If%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan("id")%>">
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
<B><%=efkan("adi")%></B></A>
</TD>


<td align="center">
<% 'ONLÝNE OLUP OLMADIÐI
sor="SELECT * FROM uyeler WHERE id = "& efkan("id") &"  "
efkan1.Open sor,Sur,1,3
zaman=datediff("n",efkan1("sontarih"),now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>" 
End If
efkan1.close
%>
</TD>


<td align="center">
<%'BU UYENÝN MESAJ VE KONU SAYISI
sor = "Select * from sorular where onay=1 and uyeid="&efkan("id")&" "  
forum.Open sor,forumbag,1,3
soruadet=forum.recordcount
forum.close
sor = "Select * from cevaplar where onay=1 and  uyeid="&efkan("id")&" "  
forum.Open sor,forumbag,1,3
cevapadet=forum.recordcount
forum.close
Response.Write soruadet &"/"&cevapadet
%>
</TD>

<td align="center"><%=efkan("hit")%></TD>
<td align="left"><%=efkan("sehir")%>&nbsp;</TD>
<td align="center"><%=efkan("sontarih")%>&nbsp;</TD>

<% If Session("efkanlogin")=True Then %>
<td align="center">
<A HREF="?part=uyegorev&gorev=bilgilerim&id=<%=efkan("id")%>">
<IMG SRC="forumimg/degistir.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="Üye Bilgilerini deðiþtir"></a>
<A HREF="?part=uyegorev&gorev=uyesil&id=<%=efkan("id")%>" onClick="return submitConfirm(this)">
<IMG SRC="forumimg/uyesil.gif" WIDTH="46" HEIGHT="17" BORDER="0" ALT=""></A>
<A HREF="?part=uyegorev&gorev=yasakla&id=<%=efkan("id")%>">
<IMG SRC="forumimg/yasakla.gif" WIDTH="42" HEIGHT="17" BORDER="0" ALT="Bu üyeye siteyi yasakla"></a>
</TD>
<% End If %>
</tr>

<% 
renk=renk + 1
efkan.movenext 
Next
efkan.close
%>
</table>

Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 20 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf ) then 
Response.Write "<B>["&y&"]</B>"
else
	  If ara<>"" then
      Response.Write "<a href='?part=uyegorev&gorev=uyeler&ara="&ara&"&shf="&y&"'>["&y&"]</a>"
      ElseIf diz<>"" then
      Response.Write "<a href=""?part=uyegorev&gorev=uyeler&diz="&diz&"&shf="&y&""">["&y&"]</a>"	  
      ElseIf harf<>"" then
      Response.Write "<a href=""?part=uyegorev&gorev=uyeler&harf="&temizle(Request("harf"))&"&shf="&y&""">["&y&"]</a>"	 	  
	  else
      Response.Write "<a href=""?part=uyegorev&gorev=uyeler&shf="&y&""">["&y&"]</a>"
      End If 
End If
say=say+1
next
End If 
End If %>






<!-- UNUTTUM BASAMAK 1 -->
<% if gorev="unuttum" then 
gkod1  =kodver2(gkod) 
%>
<B>Þifremi Unuttum Bölümü</B>
<form action="default.asp?part=uyegorev&gorev=unuttum2&SID=<%=session.sessionID%>" method="POST" >
<table width="250" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr><td align="center" valign="center" width="100%" >
<BR>
<B><FONT size=3 COLOR="red"><%=gkod1%></FONT></B><BR>
<B>*Güv.Kodu:</B><BR><input type="text" name="gkod1" size="10" maxlength="5" onkeypress="return SayiKontrol(this);"><BR>
<P><B>*Email Adresim</B><BR><input type="text" name="email" size="50" maxlength="50"> 
<P><input type="submit" value="Þifrem"></p>
<P><A HREF="default.asp?part=uyegorev&gorev=uyeol"><B>Üye Ol</B></A><BR>
<BR></TD></tr>
</table>
<% 
End If 

if gorev="unuttum2" then 
Response.Buffer = True 

'GÜVENLÝK KODU KONTROL
if  temizle(Request.Form("gkod1")) <> trim(session("gkodu2")) then
Response.Write "<BR><BR><BR><center>Güvenlik kodu yazýlmamýþ veya yanlýþ <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

formemail = temizle(Request.Form("email"))

sor="select * from uyeler where  email='"&formemail&"' and onay=1 "
efkan.Open sor,Sur,1,3

if efkan.eof or efkan.bof then
Response.Write "<BR><BR><BR><center>Verdiðiniz Bilgilerde üyemiz yok<br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
efkan.close
else

session("gkodu2")=""
'EMAÝL GÖNDERÝYOR
emesaj = "<font face=""Tahoma"" size=""2""><CENTER><b>Sayýn "&efkan("adi")
emesaj = emesaj & "<P><B>Kullanýcý adýnýz ve þifreniz aþaðýda belirtilmiþtir" 
emesaj = emesaj & "<P>Kullanýcý Adýnýz: " & efkan("kadi")
emesaj = emesaj & "<BR>Þifreniz : " & decode(efkan("sifre"))
emesaj = emesaj & "<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"'>&nbsp;Siteye gitmek için týklayýnýz</A> "
email          =efkan("email")
konu          ="Unutulan Þifre Bildirimi"
emesaj       =emesaj

call emailgonder(email,konu,emesaj)

efkan.close

Response.Write "<BR><BR><BR><BR><b>Þifreniz email adresinize gönderildi...</b><br>"
Response.Write "<b>Lütfen Emaillerinizi kontrol ediniz...</b><br>"
Response.Write "<b>Þimdi Ana Sayfaya yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='3; URL=default.asp'>" 
End If
End If
%>


<!-- BÝLGÝLERÝM -->
<% 
if gorev="bilgilerim" then 
id=kontrol(Request.QueryString("id"))

If id<>"" and Session("efkanlogin")=True then 
sor="select * from uyeler where id ="&id&"    "  'ÜYENÝN BÝLGÝLERÝ
ElseIf id="" And  Session("uyelogin")=True Then
sor="select * from uyeler where id ="&Session("uyeid")&"    "  'ÜYENÝN BÝLGÝLERÝ
else
hataver("Bu iþlem için yetkiniz yok")
Response.End
End If 
efkan.Open sor,Sur,1,3
%>

Sayýn  <B><%=efkan("adi")%></B> <BR> Bu ekranda bilgilerinizi gorebilir ve deðiþiklik yapabilirsiniz.
<P>
<B>BÝLGÝLERÝM & GÜNCELLE</B>
<form name="formcevap" onsubmit="return formCheck(this);" action="default.asp?part=uyegorev&gorev=guncelle&id=<%=id%>" method="POST" >

<table background="" width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="4">

<tr><td width="30%">Avatar Seçiniz</td><td width="70%">
<%sor = "Select * from uyeresim where uyeid = "&efkan("id")&" " 
efkan1.Open sor,Sur,1,3
resimadet=efkan1.recordcount

if efkan1.eof or efkan1.bof Then
efkan1.close

sor="SELECT * FROM avatar  "
efkan1.Open sor,Sur,1,3%>
<select name="avatar" size="9" class="form" onChange="(resim.src = avatar.options[avatar.selectedIndex].value)">
<option value="avatar/<%=efkan("avatar")%>" selected><%=efkan("avatar")%></option>
<% do while not efkan1.eof  %>
<option value="avatar/<%=efkan1("avatar")%>"><%=efkan1("avatar")%></option>
<% efkan1.movenext 
loop %>
</select>
<img src="avatar/<%=efkan("avatar")%>" width="150" height="110" name="resim">
<%
efkan1.close
Else
Response.Write "<CENTER><B>Resimlerim</B></CENTER><P>"
Response.Write "<CENTER>Resimleri büyültmek için resime týklayýnýz</CENTER><BR>"
for i=1 to uyeresimadet
if efkan1.eof then exit for%>
<A HREF="uyeler/<%=efkan1("uyeresim")%>" target="_blank">
<IMG SRC="uyeler/<%=efkan1("uyeresim")%>" WIDTH="50" HEIGHT="50" BORDER="0" ALT=""></A><BR>
<A HREF="?part=uyegorev&gorev=resimsil&id=<%=efkan1("id")%>"><B>Bu resimi sil</B></A>
<BR>

<%
efkan1.movenext
Next
efkan1.close%>
<!-- ÜYE RESÝMÝ VARSA ÇAKTIRMA AVATARI AYNEN ÝADE ET -->
<input name="avatar" size="" type="hidden"  value="<%=efkan("avatar")%>">
<%End If%>
<BR>
<!-- KAYITLI RESÝM ADETÝ VERÝLEN HAK  RESÝM YUKLEME SERBESTMÝ-->
<% if uyeresimadet >  resimadet And fotoyukleme=1 Then %>
<a href="fsoresim.asp" onClick="window.name='ana'; window.open('fsoresim.asp','new', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no, resizable=no,copyhistory=no,width=300,height=200'); return false;">Resim Yükle</a>
<% End If %>
</td></tr>

<!-- ADMÝNSE YAYINLA YAYINLAMA SEÇENEÐÝ -->
<% If Session("efkanlogin")=True Then %>
<tr><td>Üye Onayý</td><td>
*<select NAME="onay">
<option value="<%=efkan("onay")%>" selected><%=efkan("onay")%></option>
<option value="1">1</option>
<option value="0">0</option>
</select>Eðer 1 se üyelik aktiftir 0 veya baþka deðerler onaysýz veya akitvasyondan gelmeyendir</td></tr>
<% End If %>

<tr><td>Kullanýcý Adý</td><td>
*<input name="kadi" size="20" maxlength="20" value="<%=efkan("kadi")%>"></td></tr>
<tr><td>Þifreniz</td><td>
*<input type="text" name="sifre" size="20" maxlength="20" value="<%=decode(efkan("sifre"))%>"></td></tr>

<tr><td>Adýnýz</td><td>*<input type="text" name="adi" size="50" maxlength="50"  value="<%=efkan("adi")%>"></td></tr>

<tr><td>Email Adresiniz</td><td>
<% If Session("efkanlogin")=True Then %>
*<input type="text" name="email" size="50" maxlength="50"  value="<%=efkan("email")%>">
<% else %>
*<input type="text" name="email" size="50" maxlength="50"  value="<%=efkan("email")%>" readonly>
<% End If %>
</td></tr>

<tr><td>Web Sayfanýz</td><td><input type="text" name="url" size="50" maxlength="50" value="<%=efkan("url")%>"></td></tr>

<tr><td>Msn Messenger</td><td><input type="text" name="msn" size="50" maxlength="50" value="<%=efkan("msn")%>"></td></tr>

<tr><td>Yahoo</td><td><input type="text" name="yahoo" size="50" maxlength="50" value="<%=efkan("yahoo")%>"></td></tr>
<tr><td>Icq</td><td>
<input type="text" name="icq" size="50" maxlength="50" onkeypress="return SayiKontrol(this);" value="<%=efkan("icq")%>"></td></tr>

<tr><td>Mesleðiniz</td><td>
<SELECT  name="meslek"> 
<OPTION selected value="<%=efkan("meslek")%>"><%=efkan("meslek")%></OPTION> 
<OPTION value="Çalýþmýyorum">Çalýþmýyorum</OPTION>
<OPTION value="Akademisyen, Öðretmen">Akademisyen, Öðretmen</OPTION> 
<OPTION value="Avukat">Avukat</OPTION>
<OPTION value="Bankacý">Bankacý</OPTION> 
<OPTION value="Bilgisayar, Internet">Bilgisayar, Internet</OPTION>
<OPTION value="Danýþman">Danýþman</OPTION>
<OPTION value="Doktor ">Doktor</OPTION> 
<OPTION value="Emekli ">Emekli</OPTION>
<OPTION value="Ev Hanýmý ">Ev Hanýmý</OPTION>
<OPTION value="Finasman, Muhasebe ">Finasman, Muhasebe</OPTION> 
<OPTION value="Fotoðrafçý ">Fotoðrafçý</OPTION>
<OPTION value="Gazeteci ">Gazeteci</OPTION>
<OPTION value="Grafiker ">Grafiker</OPTION>
<OPTION value="Manken,Fotomodel ">Manken,Fotomodel</OPTION>
<OPTION value="Memur ">Memur</OPTION> 
<OPTION value="Mühendis ">Mühendis</OPTION>
<OPTION value="Öðrenci ">Öðrenci</OPTION> 
<OPTION value="Politikacý ">Politikacý</OPTION>
<OPTION value="Psikolog ">Psikolog</OPTION>
<OPTION value="Reklamcý ">Reklamcý</OPTION>
<OPTION value="Sanatçý ">Sanatçý</OPTION> 
<OPTION value="Satýþ, Pazarlama ">Satýþ, Pazarlama</OPTION>
<OPTION value="Serbest Meslek, Ýþ Sahibi ">Serbest Meslek, Ýþ Sahibi</OPTION>
<OPTION value="Sporcu ">Sporcu</OPTION>
<OPTION value="Teknik Eleman ">Teknik Eleman</OPTION>
<OPTION value="Üst Düzey Yönetici ">Üst Düzey Yönetici</OPTION> 
<OPTION value="Diðer ">Diðer</OPTION>
 </SELECT> 
</td></tr>



<tr><td>Yaþýnýz</td><td>
<select name="yas">
<option selected><%=efkan("yas")%></option>
<% i=16
do while i<60
i=i+1 %>
<option><%=i%></option>
<% loop %>
</select> </td></tr>


<tr><td>Þehir</td><td>
<SELECT name="sehir" class="Input" >
           <OPTION selected value="<%=efkan("sehir")%>"><%=efkan("sehir")%></OPTION>
          <OPTION value="ADANA">ADANA</OPTION>
          <OPTION value="ADIYAMAN">ADIYAMAN</OPTION>
          <OPTION value="AFYON">AFYON</OPTION>
          <OPTION value="AÐRI">AÐRI</OPTION>
          <OPTION value="AKSARAY">AKSARAY</OPTION>
          <OPTION value="AMASYA">AMASYA</OPTION>
          <OPTION value="ANKARA">ANKARA</OPTION>
          <OPTION value="ANTALYA">ANTALYA</OPTION>
          <OPTION value="ARDAHAN">ARDAHAN</OPTION>
          <OPTION value="ARTVÝN">ARTVÝN</OPTION>
          <OPTION value="AYDIN">AYDIN</OPTION>
          <OPTION value="BALIKESÝR">BALIKESÝR</OPTION>
          <OPTION value="BARTIN">BARTIN</OPTION>
          <OPTION value="BATMAN">BATMAN</OPTION>
          <OPTION value="BAYBURT">BAYBURT</OPTION>
          <OPTION value="BÝLECÝK">BÝLECÝK</OPTION>
          <OPTION value="BÝNGÖL">BÝNGÖL</OPTION>
          <OPTION value="BÝTLÝS">BÝTLÝS</OPTION>
          <OPTION value="BOLU">BOLU</OPTION>
          <OPTION value="BURDUR">BURDUR</OPTION>
          <OPTION value="BURSA">BURSA</OPTION>
          <OPTION value="ÇANAKKALE">ÇANAKKALE</OPTION>
          <OPTION value="ÇANKIRI">ÇANKIRI</OPTION>
          <OPTION value="ÇORUM">ÇORUM</OPTION>
          <OPTION value="DENÝZLÝ">DENÝZLÝ</OPTION>
          <OPTION value="DÝYARBAKIR">DÝYARBAKIR</OPTION>
          <OPTION value="DÜZCE">DÜZCE</OPTION>
          <OPTION value="EDÝRNE">EDÝRNE</OPTION>
          <OPTION value="ELAZIÐ">ELAZIÐ</OPTION>
          <OPTION value="ERZÝNCAN">ERZÝNCAN</OPTION>
          <OPTION value="ERZURUM">ERZURUM</OPTION>
          <OPTION value="ESKÝÞEHÝR">ESKÝÞEHÝR</OPTION>
          <OPTION value="GAZÝANTEP">GAZÝANTEP</OPTION>
          <OPTION value="GÝRESUN">GÝRESUN</OPTION>
          <OPTION value="GÜMÜÞHANE">GÜMÜÞHANE</OPTION>
          <OPTION value="HAKKARÝ">HAKKARÝ</OPTION>
          <OPTION value="HATAY" >HATAY</OPTION>
          <OPTION value="IÐDIR">IÐDIR</OPTION>
          <OPTION value="ISPARTA">ISPARTA</OPTION>
          <OPTION value="ÝÇEL">ÝÇEL</OPTION>
          <OPTION value="ÝSTANBUL">ÝSTANBUL</OPTION>
          <OPTION value="ÝZMÝR">ÝZMÝR</OPTION>
          <OPTION value="KAHRAMANMARAÞ">KAHRAMANMARAÞ</OPTION>
          <OPTION value="KARABÜK">KARABÜK</OPTION>
          <OPTION value="KARAMAN">KARAMAN</OPTION>
          <OPTION value="KARS">KARS</OPTION>
          <OPTION value="KASTAMONU">KASTAMONU</OPTION>
          <OPTION value="KAYSERÝ">KAYSERÝ</OPTION>
          <OPTION value="KIBRIS">KIBRIS</OPTION>
          <OPTION value="KIRIKKALE">KIRIKKALE</OPTION>
          <OPTION value="KIRKLARELÝ">KIRKLARELÝ</OPTION>
          <OPTION value="KIRÞEHÝR">KIRÞEHÝR</OPTION>
          <OPTION value="KÝLÝS">KÝLÝS</OPTION>
          <OPTION value="KOCAELÝ">KOCAELÝ</OPTION>
          <OPTION value="KONYA">KONYA</OPTION>
          <OPTION value="KÜTAHYA">KÜTAHYA</OPTION>
          <OPTION value="MALATYA">MALATYA</OPTION>
          <OPTION value="MANÝSA">MANÝSA</OPTION>
          <OPTION value="MARDÝN">MARDÝN</OPTION>
          <OPTION value="MUÐLA">MUÐLA</OPTION>
          <OPTION value="MUÞ">MUÞ</OPTION>
          <OPTION value="NEVÞEHÝR">NEVÞEHÝR</OPTION>
          <OPTION value="NÝÐDE">NÝÐDE</OPTION>
          <OPTION value="ORDU">ORDU</OPTION>
          <OPTION value="OSMANÝYE">OSMANÝYE</OPTION>
          <OPTION value="RÝZE">RÝZE</OPTION>
          <OPTION value="SAKARYA">SAKARYA</OPTION>
          <OPTION value="SAMSUN">SAMSUN</OPTION>
          <OPTION value="SÝÝRT">SÝÝRT</OPTION>
          <OPTION value="SÝNOP">SÝNOP</OPTION>
          <OPTION value="SÝVAS">SÝVAS</OPTION>
          <OPTION value="ÞANLIURFA">ÞANLIURFA</OPTION>
          <OPTION value="ÞIRNAK">ÞIRNAK</OPTION>
          <OPTION value="TEKÝRDAÐ">TEKÝRDAÐ</OPTION>
          <OPTION value="TOKAT">TOKAT</OPTION>
          <OPTION value="TRABZON">TRABZON</OPTION>
          <OPTION value="TUNCELÝ">TUNCELÝ</OPTION>
          <OPTION value="UÞAK">UÞAK</OPTION>
          <OPTION value="VAN">VAN</OPTION>
          <OPTION value="YALOVA">YALOVA</OPTION>
          <OPTION value="YOZGAT">YOZGAT</OPTION>
          <OPTION value="ZONGULDAK">ZONGULDAK</OPTION>
        </SELECT> </td></tr>


<tr><td colspan="2" align="center">
<B>Ýmzanýz </B><I>(500 Karekter)</I><P>
<!--#INCLUDE file="editor.asp"-->
<P>
<TEXTAREA  onkeyup=textKey(this.form) name="yorum" ROWS="8" COLS="80"><%=efkan("imza")%></TEXTAREA></td></tr>


<tr><td colspan="2" align="center">
<input type="submit" value="Güncelle">&nbsp;&nbsp;
<input type="reset" value="Temizle">
</td></tr></table></form>
<%efkan.close
End If  

if gorev="guncelle" then 
id = kontrol(Request.QueryString("id"))

If id<>"" and Session("efkanlogin")=True then 
sor="select * from uyeler where id ="&id&"    "  'ÜYENÝN BÝLGÝLERÝ
ElseIf id="" And  Session("uyelogin")=True Then
sor="select * from uyeler where id ="&Session("uyeid")&"    "  'ÜYENÝN BÝLGÝLERÝ
else
hataver("Bu iþlem için yetkiniz yok")
Response.End
End If 
efkan.Open sor,Sur,1,3


if request.form("kadi")="" or request.form("sifre")="" or request.form("email")="" or request.form("adi")="" then
Response.Write "<BR><BR><BR><center>Lütfen iþaretli alanlarý doldurunuz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

if emailkontrol(Request.Form ("email"))=false then
Response.Write "<BR><BR><BR><center>Geçerli Email adresi kullanýn <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

formkadi = temizle(Request.Form("kadi"))
kadisay=Len(formkadi) 
formsifre = temizle(Request.Form("sifre"))
sifresay=Len(formsifre) 

if kadisay < 4 or sifresay <4 then 
Response.Write "<BR><BR><BR><center>En az <B>4</B> karekter uzunluðunda kullanýcý adý ve þifre kullanýnýz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

'ÞU ANKÝ KULLANICI HARÝÇ BAÞKASINDA AYNI KADÝ VARMI
sor="select * from uyeler where kadi='"&formkadi&"'   and  id <> " & efkan("id") & "  " 

efkan1.Open sor,sur,1,3
adet=efkan1.recordcount
if adet > 0 Then
Response.Write "<BR><BR><BR><center>Bu kullanýcý adý kullanýlýyor baþka kullanýcý adý deneyiniz... <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
else
efkan1.close

efkan("kadi") = Temizle(Request.Form ("kadi"))
efkan("sifre") = encode(Temizle(Request.Form ("sifre")))
efkan("adi") = Temizle(Request.Form ("adi"))
efkan("email") = Temizle(Request.Form ("email"))
efkan("url") = Temizle(Request.Form ("url"))
efkan("meslek") = Temizle(Request.Form ("meslek"))
efkan("yas") = Temizle(Request.Form ("yas"))
efkan("sehir") = Temizle(Request.Form ("sehir"))
'efkan("soru") = Temizle(Request.Form ("soru"))
'efkan("cevap") = Temizle(Request.Form ("cevap"))
efkan("msn") = Temizle(Request.Form ("msn"))
efkan("yahoo") = Temizle(Request.Form ("yahoo"))
efkan("icq") = Temizle(Request.Form ("icq"))
avatar =Temizle(Request.Form ("avatar"))
avatar =Replace (avatar ,"avatar/","",1,-1,1)
efkan("avatar") =avatar

imza=suz(Temizle(Request.Form ("yorum")))
imza=left(imza,500)
efkan("imza") =imza
efkan.Update
If id<>"" and Session("efkanlogin")=True then 
efkan("onay") = Temizle(Request.Form ("onay"))
efkan.Update
Else
efkan("ipno") = Request.ServerVariables("REMOTE_ADDR")
efkan.Update
Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")
End If
Response.Write "<BR><BR><b>Bilgiler güncellendi anasayfaya yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
efkan.close
End If 
End If 


if gorev="cikis" then 
Response.Buffer = True 
cikis = DateAdd("n", -2, Now())   'ÇIKAN ÜYENÝN ZAMANINI GERÝ ALIYORUM KÝ ONLÝNE GÖRÜNMESÝN
sor="select * from uyeler where id="&Session("uyeid")&" "
efkan.Open sor,Sur,1,3
efkan("sontarih")=cikis
efkan.Update
efkan.close
session.ABANDON
Response.Redirect "default.asp"
End If



'5 AYDIR GÝRÝÞ YAPMAYAN ÜYELERÝ SÝL
if gorev="eskisil" then 
Response.Buffer = True
If Session("efkanlogin")=True <> True Then 
Response.Write "<script language='JavaScript'>alert('Bu alana girmeye yetkiniz yoktur...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyegorev'>"
Response.End
End If
Session.LCID = 1055
DefaultLCID = Session.LCID 
sor="select * from uyeler where admin<>1   " 
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("m",efkan("sontarih"),now)  ' 5 ay öncesi
if zaman > 5 then
sor="DELETE from uyeler WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
'sor = "DELETE  from mesaj where kimden = "&efkan("id")&"  or  "&efkan("id")&"      " 
'efkan2.Open sor,Sur,1,3
End If
efkan.movenext
Loop
Response.Write "<script language='JavaScript'>alert('5 aydýr giriþ yapmayanlar silindi...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
End If



'4 GÜNDÜR AKTÝVASYONDAN DÖNMEYENLERÝ SÝL
if gorev="aktifolmayansil" then 
Response.Buffer = True
If Session("efkanlogin")=True <> True Then 
Response.Write "<script language='JavaScript'>alert('Bu alana girmeye yetkiniz yoktur...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyegorev'>"
Response.End
End If
Session.LCID = 1055
DefaultLCID = Session.LCID 
sor="select * from uyeler where admin<>1 and onay <>1" 
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("d",efkan("sontarih"),now)  ' 4 gün öncesi
if zaman > 4 then
sor="DELETE from uyeler WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
Response.Write "<script language='JavaScript'>alert('4 gündür uyeliðini aktif etmeyenler silindi');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
End If


'RESÝM SÝL
if gorev="resimsil" then 
Response.Buffer = True
id=request.querystring("id")
sor="select * from uyeresim WHERE id = "&id&"  "
efkan.Open sor,Sur,1,3
If session("uyeid")=efkan("uyeid") Or Session("efkanlogin")=True  Then
sor="DELETE from uyeresim WHERE id = "&id&"  "
efkan1.Open sor,Sur,1,3
Response.Redirect "?part=uyegorev&gorev=bilgilerim"
Else
Response.Write "Bu iþlem için yetkiniz yok"
End If
efkan.close
End If



'ÜYE SÝL ADMÝN 
if gorev="uyesil" then 
If Session("efkanlogin")=True <> True Then 
Response.Write "<script language='JavaScript'>alert('Bu alana girmeye yetkiniz yoktur...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyegorev'>"
Response.End
End If
id=request.querystring("id")
sor="select * from uyeler WHERE id = "&id&"  "
efkan.Open sor,sur,1,3
If efkan("admin")<>1 Then
sor="DELETE from uyeler WHERE id = "&id&"  "
efkan1.Open sor,sur,1,3
Response.Redirect "?part=uyegorev&gorev=uyeler"
Else
Response.Write "Admini silemessin"
End If
End If



'3 AYLIK MESAJLARI SÝL NE OLURSA OLSUN
if gorev="mesajtemizle" then 
Response.Buffer = True
If Session("efkanlogin")=True <> True Then 
Response.Write "<script language='JavaScript'>alert('Bu alana girmeye yetkiniz yoktur...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyegorev'>"
Response.End
End If
Session.LCID = 1055
DefaultLCID = Session.LCID 
sor="select * from mesaj  " 
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("m",efkan("tarih"),now)  ' ay öncesi
if zaman > 3 then
sor="DELETE from mesaj WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
Response.Write "<script language='JavaScript'>alert('3 aylýk gelen giden kutularý silindi');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
End If




'//////// YASAKLILAR
if gorev="yasakli" then 
If Session("efkanlogin")=True <> True Then 
hataver("Bu iþlem için yetkinizyok")
Else %>
<A HREF="?part=uyegorev&gorev=yasakliekle">Yasaklý Ýp ve Email ekle</A><P>
<% sor = "Select * from yasakli order by id desc " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount
if efkan.eof or efkan.bof then
bilgiver("Yasaklanmýþ kiþi yok")
Else
shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>
<B>YASAKLI ÝP LER</B><BR>
<table width="95%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="3">
<tr>
<td width="1%"><B>id</B></td>
<td width="10%" align="center"><B>Tarih</B></td>
<td width="10%" align="center"><B>Ýp</B></td>
<td width="4%" align="center"><B>Ýþlem</B></td>
</tr>
<% efkan.pagesize =50  
efkan.absolutepage = shf
sayfa = efkan.pagecount
for i=1 to efkan.pagesize
if efkan.eof then exit For
%>
<tr bgcolor="" onmouseover="this.style.backgroundColor='#CCFFFF';" onmouseout="this.style.backgroundColor='';">
<td align="center"><%=efkan("id")%></td>
<td align="center"><%=efkan("tarih")%></td>
<td align="center"><%=efkan("ip")%></td>
<td align="center">
<A HREF="?part=uyegorev&gorev=yasaklisil&id=<%=efkan("id")%>" onClick="return submitConfirm(this)">Sil</A>
</td></tr>
<%
renk=renk+1
efkan.movenext
Next
efkan.close%>
</table><P>
Sayfalar :
<%say=0
for y=1 to sayfa 
if say mod 10 = 0 then
Response.Write "<BR>"
End If
if  y=cint(shf) then 
Response.Write "<B>["&y&"]</B>"
Else
	  Response.Write "<a href='?part=uyegorev&gorev=yasakli&shf="&y&"'>["&y&"]</a>"
End If
say=say+1
next
End If
End If
End If



'////////////////////////// YASAKLI EKLE /////////////////////////////////
If gorev="yasakliekle" Then 
If Session("efkanlogin")=True <> True Then 
hataver("Bu iþleme yetkiniz yok")
else%>
<form method="POST" action="?part=uyegorev&gorev=yasakliekle">
<A HREF="?part=uyegorev&gorev=yasakli">Tüm Yasaklýlar</A>
<table width="50%" bgcolor="" bordercolor="#f5f5ff" border="0" cellspacing="0" cellpadding="3">
<tr><td width="40%">Yasaklanacak Ýp</td><td width="60%">
<input name="ip" type="text" value="" size="40"></td></tr>
<tr><td align="center" colspan="2">
<input type="submit" value=" Ekle " name="submit" > <INPUT TYPE="reset" value=" Temizle ">
</td></tr></table></form>
<% 
if request.form("ip")="" And  request.form("email")="" then
else
sor = "Select * from yasakli  " 
efkan.Open sor,Sur,1,3
efkan.AddNew
  efkan("ip")         =Trim(request.form("ip"))
  efkan("tarih")     =Now()
efkan.Update
efkan.close
Response.Redirect "default.asp?part=uyegorev&gorev=yasakli"
End If
End If
End If



'//////////////////YASAKLA
if gorev="yasakla" then 
If Session("efkanlogin")=True <> True Then 
hataver("Bu iþlem için yetkinizyok")
else
id=request.querystring("id")
sor = "Select * from uyeler WHERE id = "&id&"  "
efkan.Open sor,sur,1,3
  If efkan.eof Then
  Else
  'efkan("onay")      =0
  'efkan.update
  sor = "Select * from yasakli   "
  efkan1.Open sor,sur,1,3
  efkan1.AddNew
  efkan1("ip")         =efkan("ipno")
  efkan1("tarih")     =Now()

 efkan1.update
  efkan1.close
Response.Redirect "default.asp?part=uyegorev&gorev=yasakli"
End If
efkan.close
End If
End If


'//////////////////YASAKLI SÝL
if gorev="yasaklisil" then 
If Session("efkanlogin")=True <> True Then 
hataver("Bu iþlem için yetkinizyok")
else
id=request.querystring("id")
sor="DELETE from yasakli WHERE id = "&id&"  "
efkan.Open sor,sur,1,3
Response.Redirect "default.asp?part=uyegorev&gorev=yasakli"
End If
End If











Set efkan1=Nothing
Set efkan=Nothing
%>

<P>