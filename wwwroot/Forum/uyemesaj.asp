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

<BR><BR>

<%
Response.Buffer = True


Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage

If Session("uyelogin")=True <> True Then 
Response.Redirect ("default.asp?part=uyegorev&gorev=girisform")
Response.End
End If


gorev=request.querystring("gorev") %>



<!-- GELEN KUTUSU -->
<% 
if gorev="gelen" Then
Response.Buffer = True 
sor="select * from mesaj where kime="&Session("uyeid")&" and kimesildi=0  order by id desc " 
'ÜYENÝN MESAJLARINA BAK
efkan.Open sor,Sur,1,3
gelen=efkan.recordcount  'Gelen Kutusu
%>
<IMG SRC="images/gelen.gif" WIDTH="40" HEIGHT="40" BORDER=0 ALT="">
<BR><B>Gelen Kutusu</B><P>
<B>Sayýn <%=Session("kadi")%> Gelen Kutusunda <%=gelen%> adet mesajýnýz var</B><P>

<table background="" width="98%" bgcolor="" bordercolor="#CCFFFF" border="1" cellspacing="0" cellpadding="0"><tr bgcolor="<%=bgcolor1%>" >
<td align=center width="20%"><B>Kimden</B></td>
<td align=center width="40%"><B>Konu</B></td>
<td align=center width="20%"><B>Tarih</B></td>
<td align=center width="15%" ><B>Gorev</B></td>
</tr>
<% do while not efkan.eof 
sor="select * from uyeler where id="&efkan("kimden")&" "   'GÖNDERENÝN KÝM OLDUÐUNU BUL 
efkan1.Open sor,Sur,1,3 %>
<tr><td >
<% If efkan1.eof Then
Response.Write "Bu üye silindi"
else%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan1("id")%>"><%=efkan1("kadi")%></A>
<%End If%></td>
<td>
<A HREF="default.asp?part=uyemesaj&gorev=oku&id=<%=efkan("id")%>">
<% if  efkan("okundu")=0 then %>
&nbsp;<B><%=efkan("konu")%> </B> &lt;Okunmadý&gt;
<% else %>
&nbsp;<%=efkan("konu")%>
<% End If%>
</A>
</td>
<td align=right>&nbsp;<%=efkan("tarih")%></td>
<td align=center>

<IMG SRC="images/yaz.gif" WIDTH="15"  BORDER=0 ALT="Cevapla">
</A>
&nbsp;
<A HREF="default.asp?part=uyemesaj&gorev=sil&id=<%=efkan("id")%>&kim=kimesildi">
<IMG SRC="images/del.gif" WIDTH="20" BORDER=0 ALT="Mesajý Sil"></A>
</td></tr>
<% 
efkan1.close
efkan.movenext 
loop 
efkan.close%>
</table>
<%End If %>



<!-- GÝDEN KUTUSU -->
<% if gorev="giden" then 
Response.Buffer = True 
sor="select * from mesaj where kimden="&Session("uyeid")&"  and kimdensildi=0  order by id desc"  'ÜYENÝN MESAJLARINA BAK
efkan.Open sor,Sur,1,3
giden=efkan.recordcount  'Gelen Kutusu
%>
<IMG SRC="images/giden.gif" WIDTH="40" HEIGHT="40" BORDER=0 ALT=""><BR>
<B>Giden Kutusu</B><P>
<B>Sayýn <%=Session("kadi")%> Giden Kutusunda <%=giden%> adet mesajýnýz var</B><P>
<table background="" width="98%" bgcolor="" bordercolor="#CCFFFF" border="1" cellspacing="0" cellpadding="0">
<tr bgcolor="<%=bgcolor1%>" >
<td align=center width="20%"><B>Kime</B></td>
<td align=center width="40%"><B>Konu</B></td>
<td align=center width="20%"><B>Tarih</B></td>
<td align=center width="5%"><B>Gorev</B></td>
</tr>
<% do while not efkan.eof 
sor="select * from uyeler where id="& efkan("kime")& " "   'KÝME GÖNDERDÝM
efkan1.Open sor,Sur,1,3
%>
<tr><td>
<%
If efkan1.eof Then
Response.Write "Bu üye silindi"
elseif efkan("herkese")=1 then     'kime gönderdiðini öðren 
Response.Write "Tüm Üyelere"
else%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan1("id")%>"><%=efkan1("kadi")%></A>
<%End If%>
</td>
<td>
<A HREF="default.asp?part=uyemesaj&gorev=oku&id=<%=efkan("id")%>&kim=kimdenokuyor">
&nbsp;<%=efkan("konu")%></a></td>

<td align=right>&nbsp;<%=efkan("tarih")%></td>
<td align=center>
<A HREF="default.asp?part=uyemesaj&gorev=sil&id=<%=efkan("id")%>&kim=kimdensildi">
<IMG SRC="images/del.gif" WIDTH="20" BORDER=0 ALT="Mesajý Sil"></A>
</td></tr>
<% 
efkan1.close
efkan.movenext 
loop 
efkan.close%>
</table>
<%End If 

if gorev="oku" then
Response.Buffer = True 
id= request.querystring("id")
kim= request.querystring("kim")
sor = "Select * from mesaj where id = "&id&" " 
efkan.Open sor,Sur,1,3

if kim="" then   ' kiþi giden kutusunu okuduðu zaman okundu yapmasýn diye 
efkan("okundu")= 1   'mesajý okundu yapýyorum
efkan.update
End If
%>

<IMG SRC="images/oku.gif" WIDTH="30" HEIGHT="30" BORDER=0 ALT=""><BR>
<B>Mesaj Okunuyor</B>
<P>
<table background="" width="80%" bgcolor="" bordercolor="#CCFFFF" border="1" cellspacing="0" cellpadding="5">
<tr>
<td width="20%"><B>Kimden</B></td><td width="80%">
<% 
sor="select * from uyeler where id="&efkan("kimden")&" "   'GÖNDERENÝN KÝM OLDUÐUNU BUL 
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
Response.Write "Bu üye silindi"
else%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan1("id")%>"><%=efkan1("kadi")%></A>
<%if efkan1("id")<>session("uyeid") Then 'gönderen ve okuyan aynýysa cevapla koyma%>
&nbsp;&nbsp;&nbsp;
<A HREF="default.asp?part=uyemesaj&gorev=yaz&id=<%=efkan1("id")%>&kime=<%=efkan1("kadi")%>">
<B>Cevapla</B></a>
<% 
End If 
End If 
%>
</td></tr>

<tr><td width="20%"><B>Kime</B></td><td width="80%">
<% 
efkan1.close
sor="select * from uyeler where id="&efkan("kime")&" "   'GÖNDERÝLEN KÝM OLDUÐUNU BUL 
efkan1.Open sor,Sur,1,3
If efkan1.eof Then
Response.Write "Bu üye silindi"
else%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan1("id")%>"><%=efkan1("kadi")%></A>
<% End If %>
</td></tr>

<tr><td width="20%"><B>Konu</B></td><td width="80%"><%=efkan("konu")%></td></tr>
<tr><td><B>Tarih</B></td><td>&nbsp;<%=efkan("tarih")%></td></tr>
<tr height=100><td><B>Mesaj</B></td><td><%=efkan("mesaj")%></td></tr>
</tr>
</table>
<% 
efkan1.close
efkan.close%>
<%End If %>





<% if gorev="sil" then

Response.Buffer = True 
If Session("uyelogin")=True <> True Then 
'Response.Write "<b>Bu Bölüm Üyelere Açýktýr Þimdi Yönlendiriliyorsunuz...</b><br>"
'Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyegorev&gorev=girisform'>" 
Response.Redirect ("default.asp?part=uyegorev&gorev=girisform")
Response.End
End If


id= request.querystring("id")
kim= request.querystring("kim")

sor="Select * from  mesaj WHERE id="&id&" "
efkan.Open sor,Sur,1,3

if  efkan("kimesildi")=1 and kim="kimdensildi"  then   'Gönderilen sildi ve gönderen silmek istiyorsa yani giden kutusundan siliniyorsa
efkan.close
sor = "DELETE from mesaj WHERE id="&id&""
efkan.Open sor,Sur,1,3
Response.Write "<script language='JavaScript'>alert('Mesaj Silindi...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyemesaj&gorev=giden'>"

elseif efkan("kimdensildi")=1 and kim="kimesildi"  then   'Gönderen sildi  alýcý silmek istiyorsa yani gelen kutusunda siliniyorsa
efkan.close
sor = "DELETE from mesaj WHERE id="&id&""
efkan.Open sor,Sur,1,3
Response.Write "<script language='JavaScript'>alert('Mesaj Silindi...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyemesaj&gorev=gelen'>"

elseif efkan("kimesildi")=0 and kim="kimdensildi"  then   'Gönderen silmek istiyorsa ama alýcý silmediyse
efkan("kimdensildi")=1
efkan.update
Response.Write "<script language='JavaScript'>alert('Mesaj Silindi...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyemesaj&gorev=giden'>"
efkan.close

elseif efkan("kimdensildi")=0 and kim="kimesildi"  then   'alýcý silmek istiyorsa ama gönderen silmedi ise
efkan("kimesildi")=1
efkan.update
Response.Write "<script language='JavaScript'>alert('Mesaj Silindi...');</script>"
Response.Write "<meta http-equiv='Refresh' content='0; URL=default.asp?part=uyemesaj&gorev=gelen'>"
efkan.close
End If 
End If 

 
if gorev="yaz" then 
Response.Buffer = True 

id= request.querystring("id")
kime= request.querystring("kime")
konu= request.querystring("konu")

gkod1  =kodver2(gkod)

%>

<B>Üyeye Mesaj Gönder</B><P>
<form action="default.asp?part=uyemesaj&gorev=gonder" method="POST" >
<table background="" width="60%" bgcolor="" bordercolor="#330000" border="0" cellspacing="0" cellpadding="5">
<tr><td width="30%">Güvenlik Kodu</td><td width="70%">
*<FONT SIZE="2" COLOR="red"><B><%=gkod1%></B></FONT>
<input type="text" name="gkod1" size="5" maxlength="5" onkeypress="return SayiKontrol(this);">
</td></tr>

<tr><td>Kimden</td><td>*<input value="<%=Session ("kadi")%>"	size="15" readonly></td></tr>

<tr><td>Kime</td><td>
<% sor="select * from uyeler where onay=1 order by kadi asc"
efkan.Open sor,sur,1,3%>
*<SELECT NAME="kime">
<option value="<%=id%>"><%=kime%></option>
<%do while not efkan.eof  %>
<option value="<%=efkan("id")%>"><%=efkan("kadi")%></option>
<% efkan.movenext 
loop 
'EÐER YÖNETÝCÝ ÝSE HERKESE GÖNDEREBÝLSÝN
If Session("efkanlogin") = True Then %>
<option value="0">Herkese</option>
<%End If
efkan.close%>
</SELECT></td></tr>
<tr><td>Konu</td><td>*<input name="konu" value="<%=konu%>"	size="60" maxlength="60"></td></tr>
<tr><td>Mesajýnýz</td><td>
*<TEXTAREA NAME="mesaj"  ROWS="7" COLS="60" ></TEXTAREA>
</td></tr>
<tr><td colspan="2" align="center">
<input type="hidden" name="tarih" size="30"   value="<%=(Date)%>">
<input type="submit" value="Gönder">&nbsp;&nbsp;
<input type="reset" value="Temizle">
</td></tr></table></form>
<%End If

if gorev="gonder" then 
'GÜVENLÝK KODU KONTROL
if  temizle(Request.Form("gkod1")) <> trim(session("gkodu2")) Then
Response.Write  "<BR><BR><BR><center>Güvenlik kodu yazýlmamýþ veya yanlýþ <P>Lütfen <a href=""javascript:history.back(1)""><B>&lt;&lt;geri</B></a> gidip tekrar deneyiniz"
Response.End
End If

if request.form("kime")="" or request.form("konu")="" or request.form("mesaj")=""  then
Response.Write "<BR><BR><BR><center>Lütfen tüm alanlarý doldurunuz... <br> <a href=""javascript:history.back(1)""><B>&lt;&lt;Geri</B></a><BR> gidip tekrar deneyiniz"
Response.End
End If


session("gkodu2")=""

if  request.form("kime") <>0  then    ' TEK KÝÞÝYE MESAJ GÖNDERDÝM
sor="select * from mesaj"
efkan.Open sor,sur,1,3
efkan.AddNew
efkan("kimden") =Session ("uyeid")
efkan("kime") = Temizle(Request.Form ("kime"))
efkan("konu") = suz(Temizle(Request.Form ("konu")))
mesaj=suz(Temizle(Request.Form ("mesaj")))
mesaj=left(mesaj,1000)
efkan("mesaj") =mesaj
'efkan("tarih") = Request.Form ("tarih")
efkan("tarih") = Now()
efkan.Update

' ALICIYA MESAJ VAR BÝLGÝSÝ GÖNDERÝYORUM
sor="select * from uyeler where id="& efkan("kime")& " "   'KÝME OLDUÐUNU BUL 
efkan1.Open sor,Sur,1,3
efkan.close

emesaj = "Sayýn " & efkan1("kadi") & "<P> " 
emesaj = emesaj & "Size mesaj býrakýldý mesajý okumak için  <A HREF='"&websayfam&"'><B>týklayýnýz..</B></A> "
email          =efkan1("email")
konu          ="Size mesaj býrakýldý mesajýnýz var"
emesaj       =emesaj

call emailgonder(email,konu,emesaj)

Response.Write "<BR><BR><BR><b>Mesajýnýz kaydedildi ve email adresine gönderildi....</b><br>"
Response.Write "<b>Þimdi giden kutusuna yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='1; URL=default.asp?part=uyemesaj&gorev=giden'>"
efkan1.close


else   'EÐER TEK KÝÞÝYE DEÐÝLSE  YANÝ KÝME  0  ÝSE

' HERKESE GÖNDER BÖLÜMÜ
sor="select * from uyeler"
efkan1.Open sor,sur,1,3
kimdensildi=0  ' HERKESE GÖNDERÝLEN MESAJDAN BÝR TANESÝ GÝDEN KUTUSUNDA GÖRÜNSÜN
do while not efkan1.eof
sor="select * from mesaj"
efkan.Open sor,sur,1,3
efkan.AddNew
efkan("kimden") =Session ("uyeid")
efkan("kimdensildi") =kimdensildi
efkan("kime") = efkan1("id")
efkan("herkese") = 1
efkan("konu") = suz(Temizle(Request.Form ("konu")))
mesaj=suz(Temizle(Request.Form ("mesaj")))
mesaj=left(mesaj,300)
efkan("mesaj") =mesaj
'efkan("tarih") = Request.Form ("tarih")
efkan("tarih") = Now()
efkan.Update
kimdensildi=1    ' HERKESE GÖNDERÝLEN MESAJDAN DÝÐERLERÝ GÝDEN KUTUMDA OLMASIN
' TÜM ALICILARA  MESAJ VAR BÝLGÝSÝ GÖNDERÝYORUM

emesaj = "Sayýn " & efkan1("kadi") & "<P> " 
emesaj = emesaj & "Size mesaj býrakýldý mesajý okumak için  <A HREF='"&websayfam&"'><B>týklayýnýz..</B></A> "

email          =efkan1("email")
konu          ="Size mesaj býrakýldý mesajýnýz var"
emesaj       =emesaj

call emailgonder(email,konu,emesaj)

efkan.close
efkan1.movenext 
loop 
efkan1.close
Response.Write "<BR><BR><BR><b>Mesajýnýz Tüm Üyelere gönderilmiþtir....</b><br>"
Response.Write "<b>Þimdi giden kutusuna yönlendiriliyorsunuz</b><br>"
Response.Write "<meta http-equiv='Refresh' content='1; URL=default.asp?part=uyemesaj&gorev=giden'>"
End If 

End If 
Set efkan1=Nothing
Set efkan=Nothing
%>