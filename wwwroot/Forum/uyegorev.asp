<%@CODEPAGE="1254" LANGUAGE="VbScript" LCID="1055"%>

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
KodEkle = prompt(TamYazi + "\n�rnek : http://www.makineteknik.com", YaziEkle);
if ((KodEkle != null) && (KodEkle != "")){
document.formcevap.yorum.value += "[" + Kod + "#" + KodEkle + "#]" + KodEkle + "[/"+ Kod + "] ";
}
}
document.formcevap.yorum.focus();
}

function email(Kod,TamYazi, YaziEkle) {
if (Kod != "") {
KodEkle = prompt(TamYazi + "\n�rnek : info@makineteknik.com", YaziEkle);
if ((KodEkle != null) && (KodEkle != "")){
document.formcevap.yorum.value += "[" + Kod + "#" + KodEkle + "#]" + KodEkle + "[/"+ Kod + "] ";
}
}
document.formcevap.yorum.focus();
}

function resimekle(Kod,TamYazi, YaziEkle) {
if (Kod != "") {
KodEkle = prompt(TamYazi + "\n�rnek : http://www.makineteknik.com/banner.gif", YaziEkle);
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
Response.Write "<script language='JavaScript'>alert('Bu siteye �ye giri�i yapman�z yasaklanm��t�r');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>" 
End If
efkan.close


gel = Request.ServerVariables("HTTP_REFERER")

gorev=request.querystring("gorev") 



'AKT�FLE�T�RME EMA�L�M GELMED�
if gorev="emailgelmedi" then 
If Request.ServerVariables("CONTENT_LENGTH")=0 Then %>
<B>�YEL�K AKT�FLE�T�RME EMA�L�M GELMED�</B>
<form action="default.asp?part=uyegorev&gorev=emailgelmedi" method="POST" >
<table width="400" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="5">
<tr><td align="" width="35%"><B>Email Adresiniz</B></td>
<td align="" width="65%">*<input name="email" size="30" maxlength="50"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="Tamam">
</td></tr></table></form>
<% 
Else
If Request.Form ("email")="" Then
hataver("L�tfen �ye olurken kulland���n�z email adresini yaz�n�z.")
Else
formemail =temizle(Request.Form ("email"))
sor="select * from uyeler where email='"&formemail&"'  "
efkan.Open sor,Sur,1,3
if efkan.eof or efkan.bof Then
hataver("Bu emaile ait bir kay�t bulunamad�")
ElseIf  efkan("onay")=1 Then
bilgiver("�yeli�iniz zaten aktif")
ElseIf efkan("onay")<>1 Then
emesaj = "Say�n " &efkan("adi")& " " & Now() & "<BR>"
emesaj =emesaj & websayfam & " Sitemize yapt���n�z �yelik ba�vurusunun  aktif olabilmesi i�in verilen  linke t�klay�n�z.<BR><B>4 g�n</B> sonunda �yeli�inizi aktifle�tirmezseniz.�yeli�iniz silinecektir.<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")&"'>�yeli�imi Aktif Et</A><P> "
emesaj = emesaj & " E�er verilen linkten d�nemiyorsan�z a�a��daki linki taray�c�n�z�n adres sat�r�na yap��t�rarak i�leminizi tamamlayabilirsiniz<P> "
emesaj = emesaj & websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")
email        = efkan("email")
konu        =" �yelik Aktivasyonu "
emesaj     =emesaj
call emailgonder(email,konu,emesaj)
bilgiver("Email Adresinize Aktivasyon linki g�nderildi.<BR>L�tfen emaillerinizi kontrol ediniz.<BR>4 g�n i�inde aktifle�tirilmeyen �yelikler silinecektir.")
End If
efkan.close
End If
End If
End If






'G�R�� FORM 
if gorev="girisform" then 
gkod1  =kodver2(gkod) 
%>
<form action="uyegorev.asp?gorev=kontrol" method="POST">
<table  width="300" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="5">
<tr><td colspan="2" align="center"><B>Sadece �yeler</B></td><td></tr>
<tr><td width="40%">Kullan�c� Ad�</td><td width="60%">
<input name="kadi" size="20" maxlength="10"></td></tr>
<tr><td>Parola</td><td><input type="password" name="sifre" size="20" maxlength="10"></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="�ye Giri�i"></td></tr>
<tr><td colspan="2" align="center">
<A HREF="?part=uyegorev&gorev=uyeol">�ye Olaca��m</A> |
<A HREF="?part=uyegorev&gorev=unuttum">�ifremi Unuttum</A></td></tr>
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
Response.Write "<script language='JavaScript'>alert('Kullan�c� ad� veya �ifre yanl�� l�tfen tekrar deneyin...');</script>"
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

'Y�NET�C� �SE Y�NET�C� SESS�ON AKT�F
sor="select * from uyeler where id = "& Session("uyeid") &"   "
efkan.Open sor,Sur,1,3
if efkan("admin") = 1  then
Session("efkanlogin") = True
End If
efkan.close

'OKUNMAYAN MESAJ VAR MI EN SON G�R�� TEN SONRA
sor="select * from mesaj where kime="&Session("uyeid")&" and okundu = 0 and kimesildi <>1 "
efkan.Open sor,Sur,1,3
mesajvar = efkan.recordcount
if mesajvar >0  then
Response.Write "<script language='JavaScript'>alert('Okunmayan mesaj/mesajlar var..');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyemesaj&gorev=gelen'>" 
End If
efkan.close

'SON TAR�H� G�NCELLE ONL�NE OLAYI ���N
Session.LCID = 1055
DefaultLCID = Session.LCID 

sontarih=Now()
sor="select * from uyeler where id="&Session("uyeid")&" "
efkan.Open sor,Sur,1,3
efkan("sontarih")=sontarih
efkan.Update
efkan.close
'Response.Write "<BR><BR><BR><BR><b>Ho�geldiniz&nbsp;"&formkadi&"....</b><br>"
'Response.Write "<b>�imdi Ana Sayfaya y�nlendiriliyorsunuz</b><br>"
'Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>" 
Response.Redirect "default.asp"
End If





'�YE OLMA FORMU
if gorev="uyeol" then 
gkod1  =kodver2(gkod) %>
<B>�YEL�K FORMUNU L�TFEN DOLDURUNUZ...</B>

<form name="formcevap" onsubmit="return formCheck(this);" action="default.asp?part=uyegorev&gorev=uyeoltamam" method="POST" >

<iframe src="kurallar.asp" width="500" height="100" marginwidth="0" marginheight="0" hspace="0" vspace="0" frameborder="0" scrolling="yes">
 </iframe>

<table background="" width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="4">

<tr><td width="30%">Avatar Se�iniz</td><td width="70%">
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
<% if fotoyukleme=1 Then %>�ye olduktan sonra kendi resminizi y�kleyebilirsiniz...<% End If %>
</td></tr>

<tr><td>G�venlik Kodu</td><td>
<B><FONT COLOR="red"><%=gkod1%></FONT></B>
*<input type="text" name="gkod1" size="20" maxlength="20" onkeypress="return SayiKontrol(this);">
</td></tr>

<tr><td>Kullan�c� Ad�</td><td>*<input name="kadi" size="20" maxlength="20"></td></tr>
<tr><td>�ifreniz</td><td>*<input type="password" name="sifre" size="20" maxlength="20"></td></tr>
<tr><td>Ad�n�z</td><td>*<input type="text" name="adi" size="50" maxlength="50"></td></tr>
<tr><td>Email Adresiniz</td><td>*<input type="text" name="email" size="50" maxlength="50"></td></tr>
<tr><td>Web Sayfan�z</td><td><input type="text" name="url" size="50" maxlength="50" value="http://"></td></tr>


<tr><td>Msn Messenger</td><td><input type="text" name="msn" size="50" maxlength="50"></td></tr>
<tr><td>Yahoo</td><td><input type="text" name="yahoo" size="50" maxlength="50"></td></tr>
<tr><td>Icq</td><td>
<input type="text" name="icq" size="50" maxlength="50" onkeypress="return SayiKontrol(this);"></td></tr>

<tr><td>Mesle�iniz</td><td>
<SELECT  name="meslek"> 
<OPTION selected value="">Mesle�inizi Se�in</OPTION> 
<OPTION value="�al��m�yorum">�al��m�yorum</OPTION>
<OPTION value="Akademisyen, ��retmen">Akademisyen, ��retmen</OPTION> 
<OPTION value="Avukat">Avukat</OPTION>
<OPTION value="Bankac�">Bankac�</OPTION> 
<OPTION value="Bilgisayar, Internet">Bilgisayar, Internet</OPTION>
<OPTION value="Dan��man">Dan��man</OPTION>
<OPTION value="Doktor ">Doktor</OPTION> 
<OPTION value="Emekli ">Emekli</OPTION>
<OPTION value="Ev Han�m� ">Ev Han�m�</OPTION>
<OPTION value="Finasman, Muhasebe ">Finasman, Muhasebe</OPTION> 
<OPTION value="Foto�raf�� ">Foto�raf��</OPTION>
<OPTION value="Gazeteci ">Gazeteci</OPTION>
<OPTION value="Grafiker ">Grafiker</OPTION>
<OPTION value="Manken,Fotomodel ">Manken,Fotomodel</OPTION>
<OPTION value="Memur ">Memur</OPTION> 
<OPTION value="M�hendis ">M�hendis</OPTION>
<OPTION value="��renci ">��renci</OPTION> 
<OPTION value="Politikac� ">Politikac�</OPTION>
<OPTION value="Psikolog ">Psikolog</OPTION>
<OPTION value="Reklamc� ">Reklamc�</OPTION>
<OPTION value="Sanat�� ">Sanat��</OPTION> 
<OPTION value="Sat��, Pazarlama ">Sat��, Pazarlama</OPTION>
<OPTION value="Serbest Meslek, �� Sahibi ">Serbest Meslek, �� Sahibi</OPTION>
<OPTION value="Sporcu ">Sporcu</OPTION>
<OPTION value="Teknik Eleman ">Teknik Eleman</OPTION>
<OPTION value="�st D�zey Y�netici ">�st D�zey Y�netici</OPTION> 
<OPTION value="Di�er ">Di�er</OPTION>
 </SELECT> 
</td></tr>


<tr><td>Ya��n�z</td><td>
<select name="yas">
<option selected>16</option>
<% i=16
do while i<60
i=i+1 %>
<option><%=i%></option>
<% loop %>
</select> 
</td></tr>

<tr><td>�ehir</td><td>
<SELECT name="sehir" class="Input" >
          <OPTION value="ADANA">ADANA</OPTION>
          <OPTION value="ADIYAMAN">ADIYAMAN</OPTION>
          <OPTION value="AFYON">AFYON</OPTION>
          <OPTION value="A�RI">A�RI</OPTION>
          <OPTION value="AKSARAY">AKSARAY</OPTION>
          <OPTION value="AMASYA">AMASYA</OPTION>
          <OPTION value="ANKARA">ANKARA</OPTION>
          <OPTION value="ANTALYA">ANTALYA</OPTION>
          <OPTION value="ARDAHAN">ARDAHAN</OPTION>
          <OPTION value="ARTV�N">ARTV�N</OPTION>
          <OPTION value="AYDIN">AYDIN</OPTION>
          <OPTION value="BALIKES�R">BALIKES�R</OPTION>
          <OPTION value="BARTIN">BARTIN</OPTION>
          <OPTION value="BATMAN">BATMAN</OPTION>
          <OPTION value="BAYBURT">BAYBURT</OPTION>
          <OPTION value="B�LEC�K">B�LEC�K</OPTION>
          <OPTION value="B�NG�L">B�NG�L</OPTION>
          <OPTION value="B�TL�S">B�TL�S</OPTION>
          <OPTION value="BOLU">BOLU</OPTION>
          <OPTION value="BURDUR">BURDUR</OPTION>
          <OPTION value="BURSA">BURSA</OPTION>
          <OPTION value="�ANAKKALE">�ANAKKALE</OPTION>
          <OPTION value="�ANKIRI">�ANKIRI</OPTION>
          <OPTION value="�ORUM">�ORUM</OPTION>
          <OPTION value="DEN�ZL�">DEN�ZL�</OPTION>
          <OPTION value="D�YARBAKIR">D�YARBAKIR</OPTION>
          <OPTION value="D�ZCE">D�ZCE</OPTION>
          <OPTION value="ED�RNE">ED�RNE</OPTION>
          <OPTION value="ELAZI�">ELAZI�</OPTION>
          <OPTION value="ERZ�NCAN">ERZ�NCAN</OPTION>
          <OPTION value="ERZURUM">ERZURUM</OPTION>
          <OPTION value="ESK��EH�R">ESK��EH�R</OPTION>
          <OPTION value="GAZ�ANTEP">GAZ�ANTEP</OPTION>
          <OPTION value="G�RESUN">G�RESUN</OPTION>
          <OPTION value="G�M��HANE">G�M��HANE</OPTION>
          <OPTION value="HAKKAR�">HAKKAR�</OPTION>
          <OPTION value="HATAY" >HATAY</OPTION>
          <OPTION value="I�DIR">I�DIR</OPTION>
          <OPTION value="ISPARTA">ISPARTA</OPTION>
          <OPTION value="��EL">��EL</OPTION>
          <OPTION value="�STANBUL" selected>�STANBUL</OPTION>
          <OPTION value="�ZM�R">�ZM�R</OPTION>
          <OPTION value="KAHRAMANMARA�">KAHRAMANMARA�</OPTION>
          <OPTION value="KARAB�K">KARAB�K</OPTION>
          <OPTION value="KARAMAN">KARAMAN</OPTION>
          <OPTION value="KARS">KARS</OPTION>
          <OPTION value="KASTAMONU">KASTAMONU</OPTION>
          <OPTION value="KAYSER�">KAYSER�</OPTION>
          <OPTION value="KIBRIS">KIBRIS</OPTION>
          <OPTION value="KIRIKKALE">KIRIKKALE</OPTION>
          <OPTION value="KIRKLAREL�">KIRKLAREL�</OPTION>
          <OPTION value="KIR�EH�R">KIR�EH�R</OPTION>
          <OPTION value="K�L�S">K�L�S</OPTION>
          <OPTION value="KOCAEL�">KOCAEL�</OPTION>
          <OPTION value="KONYA">KONYA</OPTION>
          <OPTION value="K�TAHYA">K�TAHYA</OPTION>
          <OPTION value="MALATYA">MALATYA</OPTION>
          <OPTION value="MAN�SA">MAN�SA</OPTION>
          <OPTION value="MARD�N">MARD�N</OPTION>
          <OPTION value="MU�LA">MU�LA</OPTION>
          <OPTION value="MU�">MU�</OPTION>
          <OPTION value="NEV�EH�R">NEV�EH�R</OPTION>
          <OPTION value="N��DE">N��DE</OPTION>
          <OPTION value="ORDU">ORDU</OPTION>
          <OPTION value="OSMAN�YE">OSMAN�YE</OPTION>
          <OPTION value="R�ZE">R�ZE</OPTION>
          <OPTION value="SAKARYA">SAKARYA</OPTION>
          <OPTION value="SAMSUN">SAMSUN</OPTION>
          <OPTION value="S��RT">S��RT</OPTION>
          <OPTION value="S�NOP">S�NOP</OPTION>
          <OPTION value="S�VAS">S�VAS</OPTION>
          <OPTION value="�ANLIURFA">�ANLIURFA</OPTION>
          <OPTION value="�IRNAK">�IRNAK</OPTION>
          <OPTION value="TEK�RDA�">TEK�RDA�</OPTION>
          <OPTION value="TOKAT">TOKAT</OPTION>
          <OPTION value="TRABZON">TRABZON</OPTION>
          <OPTION value="TUNCEL�">TUNCEL�</OPTION>
          <OPTION value="U�AK">U�AK</OPTION>
          <OPTION value="VAN">VAN</OPTION>
          <OPTION value="YALOVA">YALOVA</OPTION>
          <OPTION value="YOZGAT">YOZGAT</OPTION>
          <OPTION value="ZONGULDAK">ZONGULDAK</OPTION>
        </SELECT> 
</td></tr>

<tr><td colspan="2" align="center">
�mzan�z <I>(500 Karekter)</I><P>
<!--#INCLUDE file="editor.asp"-->
<P>
<TEXTAREA  onkeyup=textKey(this.form) name="yorum" ROWS="8" COLS="80"></TEXTAREA>
</td></tr>

<tr><td colspan="2" align="center">
<input type="hidden" name="tarih" size="30"   value="<%=(Date)%>">
<input type="submit" value="�ye Ol">&nbsp;&nbsp;
<input type="reset" value="Temizle">
</td></tr></table></form>
<%End If %>

<!-- �YE OLMA FORUMU ��LEN�YOR -->
<% 
if gorev="uyeoltamam" then 

'G�VENL�K KODU KONTROL
if  temizle(Request.Form("gkod1")) <> trim(session("gkodu2")) Then
Response.Write  "<BR><BR><BR><center>G�venlik kodu yaz�lmam�� veya yanl�� <P>L�tfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

if request.form("kadi")="" or request.form("sifre")="" or request.form("email")="" or request.form("adi")="" then
Response.Write "<BR><BR><BR><center>L�tfen i�aretli alanlar� doldurunuz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

if emailkontrol(Request.Form ("email"))=false then
Response.Write "<BR><BR><BR><center>Ge�erli Email adresi kullan�n <P>L�tfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

formkadi = temizle(Request.Form("kadi"))
kadisay=Len(formkadi) 
formsifre = temizle(Request.Form("sifre"))
sifresay=Len(formsifre) 
email=temizle(Request.Form("email"))

if kadisay < 4 or sifresay <4 then 
Response.Write "<BR><BR><BR><center>En az <B>4</B> karekter uzunlu�unda kullan�c� ad� ve �ifre kullan�n�z... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

'BU KAD� AYNI OLAN VARMI
sor="select * from uyeler where kadi='"&formkadi&"' OR email='"&email&"'   "
efkan.Open sor,sur,1,3
adet=efkan.recordcount
if adet > 0 Then
Response.Write "<BR><BR><BR><center>Bu kullan�c� ad� veya email kullan�l�yor ba�ka bir kullan�c� ad� veya email adresi deneyiniz... <P>L�tfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
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


'EMA�LLE AKT�VASYON A�IKSA 
If emaildogrulama=1 Then

minsayi = 10000 'se�ilecek say�n�n alt s�n�r�
maxsayi = 99999 'se�ilecek say�n�n �st s�n�r�
Randomize()
intRangeSize = maxsayi - minsayi + 1
sngRandomValue = intRangeSize * Rnd()
sngRandomValue = sngRandomValue + minsayi
aktifkod = Int(sngRandomValue)
efkan("onay") =aktifkod 
efkan.Update
id=efkan("id")

emesaj =websayfam & "&nbsp;Sitemize yapt���n�z �yelik ba�vurusunun  aktif olabilmesi i�in verilen  linke t�klay�n�z<P><B>4 g�n</B> sonunda �yeli�inizi aktifle�tirmezseniz.�yeli�iniz silinecektir.<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&aktifkod&"&id="&id&"'>�yeli�imi Aktif Et</A> "


emesaj = emesaj & " <P>E�er verilen linkten d�nemiyorsan�z a�a��daki linki taray�c�n�z�n adres sat�r�na yap��t�rarak i�leminizi tamamlayabilirsiniz<P> "
emesaj = emesaj & websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&aktifkod&"&id="&id


email          =efkan("email")
konu          ="�yelik Aktivasyonu"
emesaj       =emesaj
call emailgonder(email,konu,emesaj)

Response.Write "<BR><BR><BR><b>Say�n&nbsp;"&efkan("adi")&"....</b><P>"
Response.Write "<b>Email adresinize aktivasyon linki g�nderildi </b><P>"
Response.Write "<b>L�tfen emaillerinizi kontrol ediniz..... </b><P>"
Response.Write "<b>4 g�n i�inde aktifle�meyen uyelik kay�tlar� silinecektir.</b><P>"
Response.Write "<b>�imdi Ana Sayfaya y�nlendiriliyorsunuz</b><P>"
Response.Write "<meta http-equiv='Refresh' content='6; URL=default.asp'>"

'�YEL�K SERBESTSE
Else
efkan("onay") =1
efkan.Update

'�YE LOG�N OLUYOR
Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")
Response.Write "<BR><BR><BR><b>Ho�geldiniz&nbsp;"&efkan("adi")&"....</b><br>"
Response.Write "<b>�imdi Ana Sayfaya y�nlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='4; URL=default.asp'>" 
End If

'EMA�L L�STES�NE ��LEN�YOR
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



'AKT�VASYON 
if gorev="aktivasyon" then 
kod=kontrol(temizle(request.querystring("kod"))) 
id=kontrol(temizle(request.querystring("id"))) 

sor = "Select * from uyeler where onay = "&kod&" and id="&id&" " 
efkan.Open sor,Sur,1,3
if efkan.eof or efkan.bof Then
Response.Write "<B>B�yle bir Aktivasyon kodu �retilmedi</B><P>"
Response.Write "<b>�imdi Ana Sayfaya y�nlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='2; URL=default.asp'>" 
Else
efkan("onay") =1
efkan.Update
'�YE LOG�N OLUYOR
Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")
Response.Write "<BR><BR><BR><b>Ho�geldiniz&nbsp;"&efkan("adi")&"....</b><br>"
Response.Write "<b>�imdi Ana Sayfaya y�nlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='4; URL=default.asp'>" 
End If
efkan.close
End If



'///////////////// AKT�VASYONSUZLARA TOPLU AKT�VASYON EMA�L�  //////////////////////
if gorev="aktivasyonsuzlar" then 
sor = "Select * from uyeler where onay<>1  "
efkan.Open sor,sur,1,3

do while not efkan.eof 
emesaj = "Say�n " &efkan("adi")& " " & Now() & "<BR>"
emesaj =emesaj & websayfam & " Sitemize yapt���n�z �yelik ba�vurusunun  aktif olabilmesi i�in verilen  linke t�klay�n�z.<BR><B>4 g�n</B> sonunda �yeli�inizi aktifle�tirmezseniz.�yeli�iniz silinecektir.<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")&"'>�yeli�imi Aktif Et</A><P> "
emesaj = emesaj & " E�er verilen linkten d�nemiyorsan�z a�a��daki linki taray�c�n�z�n adres sat�r�na yap��t�rarak i�leminizi tamamlayabilirsiniz<P> "
emesaj = emesaj & websayfam&"/uyegorev.asp?gorev=aktivasyon&kod="&efkan("onay")&"&id="&efkan("id")

email        = efkan("email")
konu        =" �yelik Aktivasyonu "
emesaj     =emesaj
call emailgonder(email,konu,emesaj)
efkan.movenext 
loop 
efkan.close
Response.Write "�yeli�ini aktif etmeyenlere tekrar email g�nderildi"
End If



' �YE B�LG� 
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
Response.Write "<BR><BR><B>BU �YE S�L�ND�</B>" 
Response.End
End If
%>
<div align="center">
<table background="" width="80%" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="4">
<tr><td colspan="2" align="center" valign="center" width="100%" >
<B>�YE B�LG�LER�</B>
</TD></tr>


<tr><td colspan="2" align="center" valign="center" width="100%" >
<%
sor = "Select * from uyeresim where uyeid = "&id&" " 
efkan1.Open sor,Sur,1,3
if efkan1.eof or efkan1.bof Then%>
<IMG SRC="avatar/<%=efkan("avatar")%>" WIDTH="150"  BORDER="0" ALT="">
<%efkan1.close
Else
Response.Write "<CENTER>Resimleri b�y�ltmek i�in resime t�klay�n�z</CENTER><BR>"
for i=1 to uyeresimadet
if efkan1.eof then exit for%>
<A HREF="uyeler/<%=efkan1("uyeresim")%>" target="_blank">
<IMG SRC="uyeler/<%=efkan1("uyeresim")%>" WIDTH="100" HEIGHT="100" BORDER="0" ALT=""></A>

<!-- ADM�N LOG �SE RES�M S�L BUTONU -->
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
<B>Bu �yeye mesaj g�nder</B></A>

<!-- ADM�N LOG �SE UYE S�L BUTONU -->
<% If Session("efkanlogin")=True Then %>
<P>
<A HREF="?part=uyegorev&gorev=bilgilerim&id=<%=efkan("id")%>">
<IMG SRC="forumimg/degistir.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="�ye Bilgilerini de�i�tir"></a>

<A HREF="?part=uyegorev&gorev=uyesil&id=<%=efkan("id")%>">
<IMG SRC="forumimg/uyesil.gif" WIDTH="46" HEIGHT="17" BORDER="0" ALT="Bu �yeyi sil"></a>

<A HREF="?part=uyegorev&gorev=yasakla&id=<%=efkan("id")%>">
<IMG SRC="forumimg/yasakla.gif" WIDTH="42" HEIGHT="17" BORDER="0" ALT="Bu �yeye siteyi yasakla"></a>
<% End If%>

</TD></tr>

<tr><td width="30%"><B>Ad Soyad</B></td><td width="70%">
<!--E�ER Y�NET�C� �SE -->
<% If  efkan("admin") =1  then%> 
<IMG SRC="images/admin.gif" WIDTH="20" HEIGHT="20" BORDER=0 ALT="Y�netici">
<%End If%>
<%=efkan("adi")%>&nbsp;</td></tr>

<tr><td><B>�ye Adi</B></td><td><%=efkan("kadi")%></td></tr>
<tr><td><B>Web Sitesi</B></td>
<td><A HREF="<%=efkan("url")%>" target="_blank"><%=efkan("url")%>&nbsp;</A></td></tr>
<tr><td><B>Ya�</B></td><td><%=efkan("yas")%>&nbsp;</td></tr>
<tr><td><B>�ehir</B></td><td><%=efkan("sehir")%>&nbsp;</td></tr>
<tr><td><B>Meslek</B></td><td><%=efkan("meslek")%>&nbsp;</td></tr>
<tr><td><B>Msn</B></td><td><%=efkan("msn")%>&nbsp;</td></tr>
<tr><td><B>Yahoo</B></td><td><%=efkan("yahoo")%>&nbsp;</td></tr>
<tr><td><B>Icq</B></td><td><%=efkan("icq")%>&nbsp;</td></tr>
<tr><td><B>�ye Olma Tarihi</B></td><td><%=efkan("tarih")%></td></tr>
<tr><td><B>Son Giri�</B></td><td><%=efkan("sontarih")%></td></tr>
<tr><td><B>Hit</B></td><td><%=efkan("hit")%></td></tr>
<tr><td><B>Durumu</B></td><td>
<%efkan.close
sor="SELECT * FROM uyeler WHERE id = "&id&"   "
efkan.Open sor,Sur,1,3
Session.LCID = 1055
DefaultLCID = Session.LCID 
zaman=datediff("n",efkan("sontarih"),now)  ' �U AN DAN 1 DAKKA CIKAR SON TAR�H FARKI B�Y�KSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif  WIDTH=11  BORDER=0 ALT=offline>&nbsp;Online De�il" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>&nbsp;�u an Online"
End If
efkan.close
%>
</td></tr>


<!-- A�TI�I KONULAR -->
<tr><td colspan="2" align="center" valign="center" width="100%" ><B>A�t��� Son 10 Konu</B></td></tr>

<%'BU UYEN�N KONU SAYISI
sor = "Select * from sorular where onay=1 and uyeid="&id&" order by id desc "  
forum.Open sor,forumbag,1,3
soruadet=forum.recordcount
If soruadet=0 Then
Response.Write "<tr><td colspan=2>Bu �yemiz hi� konu a�mad�.</td></tr>"
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
<tr><td colspan="2" align="center" valign="center" width="100%" ><B>B�rakt��� Son 10 Mesaj</B></td></tr>
<%'BU UYEN�N MESAJ SAYISI
sor = "Select * from cevaplar where onay=1 and uyeid="&id&" order by id desc "  
forum.Open sor,forumbag,1,3
cevapadet=forum.recordcount
If cevapadet=0 Then
Response.Write "<tr><td colspan=2>Bu �yemiz hi� mesaj yazmad�.</td></tr>"
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
'/////////////////////  �YELER D�K /////////////////////////
if gorev="uyeler" Or gorev="" then 
If Session("uyelogin")=True <> True Then 
Response.Redirect ("default.asp?part=uyegorev&gorev=girisform")
Response.End
End If %>

<A HREF="?part=uyegorev&gorev=uyeler"><B>T�m �yeleri G�ster</B></A>
<table width="95%" bgcolor="" bordercolor="#CCFFFF" border="1" cellspacing="0" cellpadding="5">
<tr bgcolor=""><td align="center">
<form name="uyedok"><select name="menu" onChange="location=document.uyedok.menu.options[document.uyedok.menu.selectedIndex].value;" value="S�rala">
<option value="?part=uyegorev&gorev=uyeler">Kay�tlar� S�rala </option>
<option value="?part=uyegorev&gorev=uyeler&diz=id">Yeni Eklenenler</option>
<option value="?part=uyegorev&gorev=uyeler&diz=hit">En �ok Gelenler</option>
<option value="?part=uyegorev&gorev=uyeler&diz=eski">Eskiden Yeniye</option>
<option value="?part=uyegorev&gorev=uyeler&diz=sehir">�ehire G�re</option>
<option value="?part=uyegorev&gorev=uyeler&diz=kadi">Kullan�c� Ad�na G�re</option>
<option value="?part=uyegorev&gorev=uyeler&diz=sontarihe">Son Giri� yapanlar</option>
<% If Session("efkanlogin")=True Then %>
<option value="?part=uyegorev&gorev=uyeler&diz=onaysiz">Onay Bekleyenler</option>
<% End If %></select></td></form>

<!-- ARAMA B�L�M� -->
<FORM method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?part=uyegorev&gorev=uyeler">
<td align="center">
<input type="text" name="ara" size="20" maxlength="20">&nbsp;&nbsp;<input type="submit" value=" Ara ">
</td></FORM></tr></table>


<%
diz       =trim(Request("diz"))  'D�ZME SE�ENEKLER�
ara      =temizle(Request("ara"))



If diz="id" then
sor = "Select * from uyeler  order by id desc"
mesaj="YEN� �YELER�M�Z "
ElseIf diz="onaysiz" then
sor = "Select * from uyeler where onay<>1 order by id desc"
mesaj="ONAY BEKLEYENLER "
ElseIf diz="eski" then
sor = "Select * from uyeler  order by id asc"
mesaj="ESK� �YELER�M�Z "
ElseIf diz="sehir" then
sor = "Select * from uyeler  order by sehir asc"
mesaj="�EH�RE G�RE D�Z�LD� "
ElseIf diz="kadi" then
sor = "Select * from uyeler  order by kadi asc"
mesaj="KULLANICI ADINA G�RE D�Z�LD� "
ElseIf diz="sontarihe" then
sor = "Select * from uyeler  order by sontarih desc"
mesaj="SON G�R��E G�RE D�Z�LD�"
ElseIf diz="hit" then
sor = "Select * from uyeler order by hit desc"
mesaj="EN �OK G�R�� YAPANLAR"
ElseIf ara<>"" then
sor = "select * from uyeler WHERE onay=1 AND kadi like '%"&ara&"%' OR  onay=1 AND email like '%"&ara&"%' order by id desc "
mesaj="ARAMA SONUCU"
Else 
sor = "Select * from uyeler order by id desc"
mesaj="YEN� �YELER�M�Z "
End If

efkan.Open sor,Sur,1,3
adet=efkan.recordcount

if efkan.eof or efkan.bof then
hataver("Kay�t Bulunamad�")
else

shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>

<P><B><%=mesaj%>&nbsp;</B><BR> Toplam <%=adet%> kay�t
<table background="" width="99%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="3">
<tr bgcolor="<%=bgcolor1%>">
<td align="center" width="5%" ><B>Ad� </B></TD>
<td align="center" width="1%" ><B>Online</B></TD>
<td align="center" width="1%" ><B>Konu<BR>mesaj</B></TD>
<td align="center" width="1%" ><B>Hit</B></TD>
<td align="center" width="5%" ><B>�ehir</B></TD>
<td align="center" width="5%" ><B>Son Giri�</B></TD>

<% If Session("efkanlogin")=True Then %>
<td align="center" width="5%" ><B>��lem</B></TD>
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
<!--E�ER Y�NET�C� �SE -->
<% If  efkan("admin") =1  then%> 
<IMG SRC="images/admin.gif" WIDTH="20" HEIGHT="20" BORDER=0 ALT="Y�netici">
<%End If%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan("id")%>">
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
<B><%=efkan("adi")%></B></A>
</TD>


<td align="center">
<% 'ONL�NE OLUP OLMADI�I
sor="SELECT * FROM uyeler WHERE id = "& efkan("id") &"  "
efkan1.Open sor,Sur,1,3
zaman=datediff("n",efkan1("sontarih"),now)  ' �U AN DAN 1 DAKKA CIKAR SON TAR�H FARKI B�Y�KSE
if zaman > 1 then
Response.Write "<IMG SRC=images/off.gif WIDTH=11  BORDER=0 ALT=offline>" 
else 
Response.Write "<IMG SRC=images/onn.gif WIDTH=11  BORDER=0 ALT=online>" 
End If
efkan1.close
%>
</TD>


<td align="center">
<%'BU UYEN�N MESAJ VE KONU SAYISI
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
<IMG SRC="forumimg/degistir.gif" WIDTH="55" HEIGHT="17" BORDER="0" ALT="�ye Bilgilerini de�i�tir"></a>
<A HREF="?part=uyegorev&gorev=uyesil&id=<%=efkan("id")%>" onClick="return submitConfirm(this)">
<IMG SRC="forumimg/uyesil.gif" WIDTH="46" HEIGHT="17" BORDER="0" ALT=""></A>
<A HREF="?part=uyegorev&gorev=yasakla&id=<%=efkan("id")%>">
<IMG SRC="forumimg/yasakla.gif" WIDTH="42" HEIGHT="17" BORDER="0" ALT="Bu �yeye siteyi yasakla"></a>
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
<B>�ifremi Unuttum B�l�m�</B>
<form action="default.asp?part=uyegorev&gorev=unuttum2&SID=<%=session.sessionID%>" method="POST" >
<table width="250" bgcolor="" bordercolor="" border="0" cellspacing="0" cellpadding="0">
<tr><td align="center" valign="center" width="100%" >
<BR>
<B><FONT size=3 COLOR="red"><%=gkod1%></FONT></B><BR>
<B>*G�v.Kodu:</B><BR><input type="text" name="gkod1" size="10" maxlength="5" onkeypress="return SayiKontrol(this);"><BR>
<P><B>*Email Adresim</B><BR><input type="text" name="email" size="50" maxlength="50"> 
<P><input type="submit" value="�ifrem"></p>
<P><A HREF="default.asp?part=uyegorev&gorev=uyeol"><B>�ye Ol</B></A><BR>
<BR></TD></tr>
</table>
<% 
End If 

if gorev="unuttum2" then 
Response.Buffer = True 

'G�VENL�K KODU KONTROL
if  temizle(Request.Form("gkod1")) <> trim(session("gkodu2")) then
Response.Write "<BR><BR><BR><center>G�venlik kodu yaz�lmam�� veya yanl�� <P>L�tfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

formemail = temizle(Request.Form("email"))

sor="select * from uyeler where  email='"&formemail&"' and onay=1 "
efkan.Open sor,Sur,1,3

if efkan.eof or efkan.bof then
Response.Write "<BR><BR><BR><center>Verdi�iniz Bilgilerde �yemiz yok<br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
efkan.close
else

session("gkodu2")=""
'EMA�L G�NDER�YOR
emesaj = "<font face=""Tahoma"" size=""2""><CENTER><b>Say�n "&efkan("adi")
emesaj = emesaj & "<P><B>Kullan�c� ad�n�z ve �ifreniz a�a��da belirtilmi�tir" 
emesaj = emesaj & "<P>Kullan�c� Ad�n�z: " & efkan("kadi")
emesaj = emesaj & "<BR>�ifreniz : " & decode(efkan("sifre"))
emesaj = emesaj & "<P>"
emesaj = emesaj & "<A HREF='"&websayfam&"'>&nbsp;Siteye gitmek i�in t�klay�n�z</A> "
email          =efkan("email")
konu          ="Unutulan �ifre Bildirimi"
emesaj       =emesaj

call emailgonder(email,konu,emesaj)

efkan.close

Response.Write "<BR><BR><BR><BR><b>�ifreniz email adresinize g�nderildi...</b><br>"
Response.Write "<b>L�tfen Emaillerinizi kontrol ediniz...</b><br>"
Response.Write "<b>�imdi Ana Sayfaya y�nlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='3; URL=default.asp'>" 
End If
End If
%>


<!-- B�LG�LER�M -->
<% 
if gorev="bilgilerim" then 
id=kontrol(Request.QueryString("id"))

If id<>"" and Session("efkanlogin")=True then 
sor="select * from uyeler where id ="&id&"    "  '�YEN�N B�LG�LER�
ElseIf id="" And  Session("uyelogin")=True Then
sor="select * from uyeler where id ="&Session("uyeid")&"    "  '�YEN�N B�LG�LER�
else
hataver("Bu i�lem i�in yetkiniz yok")
Response.End
End If 
efkan.Open sor,Sur,1,3
%>

Say�n  <B><%=efkan("adi")%></B> <BR> Bu ekranda bilgilerinizi gorebilir ve de�i�iklik yapabilirsiniz.
<P>
<B>B�LG�LER�M & G�NCELLE</B>
<form name="formcevap" onsubmit="return formCheck(this);" action="default.asp?part=uyegorev&gorev=guncelle&id=<%=id%>" method="POST" >

<table background="" width="60%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="4">

<tr><td width="30%">Avatar Se�iniz</td><td width="70%">
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
Response.Write "<CENTER>Resimleri b�y�ltmek i�in resime t�klay�n�z</CENTER><BR>"
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
<!-- �YE RES�M� VARSA �AKTIRMA AVATARI AYNEN �ADE ET -->
<input name="avatar" size="" type="hidden"  value="<%=efkan("avatar")%>">
<%End If%>
<BR>
<!-- KAYITLI RES�M ADET� VER�LEN HAK  RES�M YUKLEME SERBESTM�-->
<% if uyeresimadet >  resimadet And fotoyukleme=1 Then %>
<a href="fsoresim.asp" onClick="window.name='ana'; window.open('fsoresim.asp','new', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no, resizable=no,copyhistory=no,width=300,height=200'); return false;">Resim Y�kle</a>
<% End If %>
</td></tr>

<!-- ADM�NSE YAYINLA YAYINLAMA SE�ENE�� -->
<% If Session("efkanlogin")=True Then %>
<tr><td>�ye Onay�</td><td>
*<select NAME="onay">
<option value="<%=efkan("onay")%>" selected><%=efkan("onay")%></option>
<option value="1">1</option>
<option value="0">0</option>
</select>E�er 1 se �yelik aktiftir 0 veya ba�ka de�erler onays�z veya akitvasyondan gelmeyendir</td></tr>
<% End If %>

<tr><td>Kullan�c� Ad�</td><td>
*<input name="kadi" size="20" maxlength="20" value="<%=efkan("kadi")%>"></td></tr>
<tr><td>�ifreniz</td><td>
*<input type="text" name="sifre" size="20" maxlength="20" value="<%=decode(efkan("sifre"))%>"></td></tr>

<tr><td>Ad�n�z</td><td>*<input type="text" name="adi" size="50" maxlength="50"  value="<%=efkan("adi")%>"></td></tr>

<tr><td>Email Adresiniz</td><td>
<% If Session("efkanlogin")=True Then %>
*<input type="text" name="email" size="50" maxlength="50"  value="<%=efkan("email")%>">
<% else %>
*<input type="text" name="email" size="50" maxlength="50"  value="<%=efkan("email")%>" readonly>
<% End If %>
</td></tr>

<tr><td>Web Sayfan�z</td><td><input type="text" name="url" size="50" maxlength="50" value="<%=efkan("url")%>"></td></tr>

<tr><td>Msn Messenger</td><td><input type="text" name="msn" size="50" maxlength="50" value="<%=efkan("msn")%>"></td></tr>

<tr><td>Yahoo</td><td><input type="text" name="yahoo" size="50" maxlength="50" value="<%=efkan("yahoo")%>"></td></tr>
<tr><td>Icq</td><td>
<input type="text" name="icq" size="50" maxlength="50" onkeypress="return SayiKontrol(this);" value="<%=efkan("icq")%>"></td></tr>

<tr><td>Mesle�iniz</td><td>
<SELECT  name="meslek"> 
<OPTION selected value="<%=efkan("meslek")%>"><%=efkan("meslek")%></OPTION> 
<OPTION value="�al��m�yorum">�al��m�yorum</OPTION>
<OPTION value="Akademisyen, ��retmen">Akademisyen, ��retmen</OPTION> 
<OPTION value="Avukat">Avukat</OPTION>
<OPTION value="Bankac�">Bankac�</OPTION> 
<OPTION value="Bilgisayar, Internet">Bilgisayar, Internet</OPTION>
<OPTION value="Dan��man">Dan��man</OPTION>
<OPTION value="Doktor ">Doktor</OPTION> 
<OPTION value="Emekli ">Emekli</OPTION>
<OPTION value="Ev Han�m� ">Ev Han�m�</OPTION>
<OPTION value="Finasman, Muhasebe ">Finasman, Muhasebe</OPTION> 
<OPTION value="Foto�raf�� ">Foto�raf��</OPTION>
<OPTION value="Gazeteci ">Gazeteci</OPTION>
<OPTION value="Grafiker ">Grafiker</OPTION>
<OPTION value="Manken,Fotomodel ">Manken,Fotomodel</OPTION>
<OPTION value="Memur ">Memur</OPTION> 
<OPTION value="M�hendis ">M�hendis</OPTION>
<OPTION value="��renci ">��renci</OPTION> 
<OPTION value="Politikac� ">Politikac�</OPTION>
<OPTION value="Psikolog ">Psikolog</OPTION>
<OPTION value="Reklamc� ">Reklamc�</OPTION>
<OPTION value="Sanat�� ">Sanat��</OPTION> 
<OPTION value="Sat��, Pazarlama ">Sat��, Pazarlama</OPTION>
<OPTION value="Serbest Meslek, �� Sahibi ">Serbest Meslek, �� Sahibi</OPTION>
<OPTION value="Sporcu ">Sporcu</OPTION>
<OPTION value="Teknik Eleman ">Teknik Eleman</OPTION>
<OPTION value="�st D�zey Y�netici ">�st D�zey Y�netici</OPTION> 
<OPTION value="Di�er ">Di�er</OPTION>
 </SELECT> 
</td></tr>



<tr><td>Ya��n�z</td><td>
<select name="yas">
<option selected><%=efkan("yas")%></option>
<% i=16
do while i<60
i=i+1 %>
<option><%=i%></option>
<% loop %>
</select> </td></tr>


<tr><td>�ehir</td><td>
<SELECT name="sehir" class="Input" >
           <OPTION selected value="<%=efkan("sehir")%>"><%=efkan("sehir")%></OPTION>
          <OPTION value="ADANA">ADANA</OPTION>
          <OPTION value="ADIYAMAN">ADIYAMAN</OPTION>
          <OPTION value="AFYON">AFYON</OPTION>
          <OPTION value="A�RI">A�RI</OPTION>
          <OPTION value="AKSARAY">AKSARAY</OPTION>
          <OPTION value="AMASYA">AMASYA</OPTION>
          <OPTION value="ANKARA">ANKARA</OPTION>
          <OPTION value="ANTALYA">ANTALYA</OPTION>
          <OPTION value="ARDAHAN">ARDAHAN</OPTION>
          <OPTION value="ARTV�N">ARTV�N</OPTION>
          <OPTION value="AYDIN">AYDIN</OPTION>
          <OPTION value="BALIKES�R">BALIKES�R</OPTION>
          <OPTION value="BARTIN">BARTIN</OPTION>
          <OPTION value="BATMAN">BATMAN</OPTION>
          <OPTION value="BAYBURT">BAYBURT</OPTION>
          <OPTION value="B�LEC�K">B�LEC�K</OPTION>
          <OPTION value="B�NG�L">B�NG�L</OPTION>
          <OPTION value="B�TL�S">B�TL�S</OPTION>
          <OPTION value="BOLU">BOLU</OPTION>
          <OPTION value="BURDUR">BURDUR</OPTION>
          <OPTION value="BURSA">BURSA</OPTION>
          <OPTION value="�ANAKKALE">�ANAKKALE</OPTION>
          <OPTION value="�ANKIRI">�ANKIRI</OPTION>
          <OPTION value="�ORUM">�ORUM</OPTION>
          <OPTION value="DEN�ZL�">DEN�ZL�</OPTION>
          <OPTION value="D�YARBAKIR">D�YARBAKIR</OPTION>
          <OPTION value="D�ZCE">D�ZCE</OPTION>
          <OPTION value="ED�RNE">ED�RNE</OPTION>
          <OPTION value="ELAZI�">ELAZI�</OPTION>
          <OPTION value="ERZ�NCAN">ERZ�NCAN</OPTION>
          <OPTION value="ERZURUM">ERZURUM</OPTION>
          <OPTION value="ESK��EH�R">ESK��EH�R</OPTION>
          <OPTION value="GAZ�ANTEP">GAZ�ANTEP</OPTION>
          <OPTION value="G�RESUN">G�RESUN</OPTION>
          <OPTION value="G�M��HANE">G�M��HANE</OPTION>
          <OPTION value="HAKKAR�">HAKKAR�</OPTION>
          <OPTION value="HATAY" >HATAY</OPTION>
          <OPTION value="I�DIR">I�DIR</OPTION>
          <OPTION value="ISPARTA">ISPARTA</OPTION>
          <OPTION value="��EL">��EL</OPTION>
          <OPTION value="�STANBUL">�STANBUL</OPTION>
          <OPTION value="�ZM�R">�ZM�R</OPTION>
          <OPTION value="KAHRAMANMARA�">KAHRAMANMARA�</OPTION>
          <OPTION value="KARAB�K">KARAB�K</OPTION>
          <OPTION value="KARAMAN">KARAMAN</OPTION>
          <OPTION value="KARS">KARS</OPTION>
          <OPTION value="KASTAMONU">KASTAMONU</OPTION>
          <OPTION value="KAYSER�">KAYSER�</OPTION>
          <OPTION value="KIBRIS">KIBRIS</OPTION>
          <OPTION value="KIRIKKALE">KIRIKKALE</OPTION>
          <OPTION value="KIRKLAREL�">KIRKLAREL�</OPTION>
          <OPTION value="KIR�EH�R">KIR�EH�R</OPTION>
          <OPTION value="K�L�S">K�L�S</OPTION>
          <OPTION value="KOCAEL�">KOCAEL�</OPTION>
          <OPTION value="KONYA">KONYA</OPTION>
          <OPTION value="K�TAHYA">K�TAHYA</OPTION>
          <OPTION value="MALATYA">MALATYA</OPTION>
          <OPTION value="MAN�SA">MAN�SA</OPTION>
          <OPTION value="MARD�N">MARD�N</OPTION>
          <OPTION value="MU�LA">MU�LA</OPTION>
          <OPTION value="MU�">MU�</OPTION>
          <OPTION value="NEV�EH�R">NEV�EH�R</OPTION>
          <OPTION value="N��DE">N��DE</OPTION>
          <OPTION value="ORDU">ORDU</OPTION>
          <OPTION value="OSMAN�YE">OSMAN�YE</OPTION>
          <OPTION value="R�ZE">R�ZE</OPTION>
          <OPTION value="SAKARYA">SAKARYA</OPTION>
          <OPTION value="SAMSUN">SAMSUN</OPTION>
          <OPTION value="S��RT">S��RT</OPTION>
          <OPTION value="S�NOP">S�NOP</OPTION>
          <OPTION value="S�VAS">S�VAS</OPTION>
          <OPTION value="�ANLIURFA">�ANLIURFA</OPTION>
          <OPTION value="�IRNAK">�IRNAK</OPTION>
          <OPTION value="TEK�RDA�">TEK�RDA�</OPTION>
          <OPTION value="TOKAT">TOKAT</OPTION>
          <OPTION value="TRABZON">TRABZON</OPTION>
          <OPTION value="TUNCEL�">TUNCEL�</OPTION>
          <OPTION value="U�AK">U�AK</OPTION>
          <OPTION value="VAN">VAN</OPTION>
          <OPTION value="YALOVA">YALOVA</OPTION>
          <OPTION value="YOZGAT">YOZGAT</OPTION>
          <OPTION value="ZONGULDAK">ZONGULDAK</OPTION>
        </SELECT> </td></tr>


<tr><td colspan="2" align="center">
<B>�mzan�z </B><I>(500 Karekter)</I><P>
<!--#INCLUDE file="editor.asp"-->
<P>
<TEXTAREA  onkeyup=textKey(this.form) name="yorum" ROWS="8" COLS="80"><%=efkan("imza")%></TEXTAREA></td></tr>


<tr><td colspan="2" align="center">
<input type="submit" value="G�ncelle">&nbsp;&nbsp;
<input type="reset" value="Temizle">
</td></tr></table></form>
<%efkan.close
End If  

if gorev="guncelle" then 
id = kontrol(Request.QueryString("id"))

If id<>"" and Session("efkanlogin")=True then 
sor="select * from uyeler where id ="&id&"    "  '�YEN�N B�LG�LER�
ElseIf id="" And  Session("uyelogin")=True Then
sor="select * from uyeler where id ="&Session("uyeid")&"    "  '�YEN�N B�LG�LER�
else
hataver("Bu i�lem i�in yetkiniz yok")
Response.End
End If 
efkan.Open sor,Sur,1,3


if request.form("kadi")="" or request.form("sifre")="" or request.form("email")="" or request.form("adi")="" then
Response.Write "<BR><BR><BR><center>L�tfen i�aretli alanlar� doldurunuz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

if emailkontrol(Request.Form ("email"))=false then
Response.Write "<BR><BR><BR><center>Ge�erli Email adresi kullan�n <P>L�tfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

formkadi = temizle(Request.Form("kadi"))
kadisay=Len(formkadi) 
formsifre = temizle(Request.Form("sifre"))
sifresay=Len(formsifre) 

if kadisay < 4 or sifresay <4 then 
Response.Write "<BR><BR><BR><center>En az <B>4</B> karekter uzunlu�unda kullan�c� ad� ve �ifre kullan�n�z... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If

'�U ANK� KULLANICI HAR�� BA�KASINDA AYNI KAD� VARMI
sor="select * from uyeler where kadi='"&formkadi&"'   and  id <> " & efkan("id") & "  " 

efkan1.Open sor,sur,1,3
adet=efkan1.recordcount
if adet > 0 Then
Response.Write "<BR><BR><BR><center>Bu kullan�c� ad� kullan�l�yor ba�ka kullan�c� ad� deneyiniz... <P>L�tfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
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
Response.Write "<BR><BR><b>Bilgiler g�ncellendi anasayfaya y�nlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
efkan.close
End If 
End If 


if gorev="cikis" then 
Response.Buffer = True 
cikis = DateAdd("n", -2, Now())   '�IKAN �YEN�N ZAMANINI GER� ALIYORUM K� ONL�NE G�R�NMES�N
sor="select * from uyeler where id="&Session("uyeid")&" "
efkan.Open sor,Sur,1,3
efkan("sontarih")=cikis
efkan.Update
efkan.close
session.ABANDON
Response.Redirect "default.asp"
End If



'5 AYDIR G�R�� YAPMAYAN �YELER� S�L
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
zaman=datediff("m",efkan("sontarih"),now)  ' 5 ay �ncesi
if zaman > 5 then
sor="DELETE from uyeler WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
'sor = "DELETE  from mesaj where kimden = "&efkan("id")&"  or  "&efkan("id")&"      " 
'efkan2.Open sor,Sur,1,3
End If
efkan.movenext
Loop
Response.Write "<script language='JavaScript'>alert('5 ayd�r giri� yapmayanlar silindi...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
End If



'4 G�ND�R AKT�VASYONDAN D�NMEYENLER� S�L
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
zaman=datediff("d",efkan("sontarih"),now)  ' 4 g�n �ncesi
if zaman > 4 then
sor="DELETE from uyeler WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
Response.Write "<script language='JavaScript'>alert('4 g�nd�r uyeli�ini aktif etmeyenler silindi');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
End If


'RES�M S�L
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
Response.Write "Bu i�lem i�in yetkiniz yok"
End If
efkan.close
End If



'�YE S�L ADM�N 
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



'3 AYLIK MESAJLARI S�L NE OLURSA OLSUN
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
zaman=datediff("m",efkan("tarih"),now)  ' ay �ncesi
if zaman > 3 then
sor="DELETE from mesaj WHERE id = "&efkan("id")&"  "
efkan1.Open sor,Sur,1,3
End If
efkan.movenext
Loop
Response.Write "<script language='JavaScript'>alert('3 ayl�k gelen giden kutular� silindi');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp'>"
End If




'//////// YASAKLILAR
if gorev="yasakli" then 
If Session("efkanlogin")=True <> True Then 
hataver("Bu i�lem i�in yetkinizyok")
Else %>
<A HREF="?part=uyegorev&gorev=yasakliekle">Yasakl� �p ve Email ekle</A><P>
<% sor = "Select * from yasakli order by id desc " 
efkan.Open sor,Sur,1,3
adet=efkan.recordcount
if efkan.eof or efkan.bof then
bilgiver("Yasaklanm�� ki�i yok")
Else
shf = Request.QueryString("shf")
if shf="" then 
shf=1
End If %>
<B>YASAKLI �P LER</B><BR>
<table width="95%" bgcolor="" bordercolor="#FFFFFF" border="0" cellspacing="0" cellpadding="3">
<tr>
<td width="1%"><B>id</B></td>
<td width="10%" align="center"><B>Tarih</B></td>
<td width="10%" align="center"><B>�p</B></td>
<td width="4%" align="center"><B>��lem</B></td>
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
hataver("Bu i�leme yetkiniz yok")
else%>
<form method="POST" action="?part=uyegorev&gorev=yasakliekle">
<A HREF="?part=uyegorev&gorev=yasakli">T�m Yasakl�lar</A>
<table width="50%" bgcolor="" bordercolor="#f5f5ff" border="0" cellspacing="0" cellpadding="3">
<tr><td width="40%">Yasaklanacak �p</td><td width="60%">
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
hataver("Bu i�lem i�in yetkinizyok")
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


'//////////////////YASAKLI S�L
if gorev="yasaklisil" then 
If Session("efkanlogin")=True <> True Then 
hataver("Bu i�lem i�in yetkinizyok")
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