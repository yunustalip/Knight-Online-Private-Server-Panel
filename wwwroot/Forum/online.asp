

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

<%
 'ÜYE VARSA ZAMAN TAZELE
If Session("uyelogin")=True   then
sontarih=Now()
sor="select * from uyeler where id="&Session("uyeid")&" "
efkan.Open sor,Sur,1,3
efkan("sontarih")=sontarih
efkan.Update

Session("uyelogin") = True
Session ("kadi") = efkan("kadi")
Session ("uyeid") = efkan("id")
Session ("adi") = efkan("adi")
Session ("email") = efkan("email")
efkan.close
End If

 
 'TOPLAM ÜYE SAYISI
sor="select * from uyeler "
efkan.Open sor,Sur,1,3
toplamuye=efkan.recordcount  'TOPLAM ÜYE SAYISI
efkan.close

 
'ONLÝNE ZÝYARETÇÝ 
zamanmiktari=1 
sor = "SELECT * FROM online"
efkan.open sor, sur, 1, 3
Do While Not efkan.eof 
zaman=datediff("n",efkan("zaman"),now)
if zaman > zamanmiktari then
sor = "DELETE FROM online WHERE  ipno = '"&efkan("ipno")&"'"
efkan1.open sor, sur, 1, 3
End If
efkan.movenext
Loop
onlineadet = efkan.RecordCount 'ONLÝNE TOPLAM ZÝYARETCÝ
efkan.Close

 
'SAYAÇ BAÞLA
Dim ip_no,site_name,zaman,site_gel
ip_no=      Request.ServerVariables("REMOTE_ADDR") 
site_ad=    Request.ServerVariables("URL") 
site_gel=    Request.ServerVariables("HTTP_REFERER") 
if site_gel="" then 
site_gel="Anasayfam"
else
uzunluk=len(site_gel)
kisa=mid(site_gel,8,uzunluk) 
bul=instr(kisa,"/")
if bul<>"0" then
site_gel=mid(site_gel,1,bul+6) 
End If
End If

zaman=      mid(now(),1,10) 'BUGUN
if session("ziyaret")<>"yes" then

'HÝT GONDEREN SÝTELERÝN TOPLAM SAYACI
sor="Select * from say_site where  site_name like '"&site_gel&"'  "
efkan.Open sor,Sur,1,3
if efkan.eof then 
efkan.AddNew
efkan("site_name")=site_gel
efkan("hit")=1
efkan("gun")=zaman
efkan.Update 
efkan.close 
else
efkan("hit")=efkan("hit")+1
efkan.Update 
efkan.close 
End If


'HÝT GONDEREN SÝTELER GUNLUK SAYACI
sor="Select * from site_gel where (gun like '"&zaman&"' and site_gel like '"&site_gel&"')"
efkan.Open sor,Sur,1,3
if efkan.eof then 
efkan.AddNew
efkan("site_gel")=site_gel
efkan("hit")=1
efkan("gun")=zaman
efkan.Update 
efkan.close 
else
efkan("hit")=efkan("hit")+1
efkan.Update 
efkan.close 
End If


'EN ÇOK ZÝYARET EDEN ÝPLER ÝP TOPLAM SAYACI
sor="Select * from say_ip where  ip_number like '"&ip_no&"' "  'ÝP ÝLKEZ GELÝYORSA
efkan.Open sor,Sur,1,3
if efkan.eof then 
efkan.AddNew
efkan("ip_number")=ip_no
efkan("hit")=1
efkan("vakit")=zaman
efkan.Update 
efkan.close 
tekil="ok"
else

if efkan("vakit") <> zaman then tekil="ok" else tekil="no" End If  'ÝP KAYITLI AMA BUGUN GELMEDÝYSE T. SAY
efkan("hit")=efkan("hit")+1
efkan("vakit")=zaman
efkan.Update 
efkan.close 
End If


'GÜNLÜK HÝTÝ 
Sor="Select * from say_hit where gun like '"&zaman&"'"
efkan.Open sor,Sur,1,3
if efkan.eof then 
efkan.AddNew
efkan("gun")=zaman
efkan("tekil")=1
efkan("cogul")=1
efkan.Update 
efkan.close 
else

if tekil="ok" then  'EGER BUGUN GELMEDÝ ÝSE
efkan("tekil")=efkan("tekil")+1
efkan("cogul")=efkan("cogul")+1
efkan.Update 
efkan.close 
else 'EÐER BUGUN ÖNCEDEN GÝRÝÞ YAPTISA
efkan("cogul")=efkan("cogul")+1
efkan.Update 
efkan.close 
End If
End If


End If

Dim gunt,gunc,topt,topc
gunt=0
gunc=0
topt=0
topc=0

Sor="Select * from say_hit"
efkan.Open sor,Sur,1,3
toplamgun=efkan.recordcount  
Do while not efkan.Eof

if efkan("gun")=zaman then
gunt=efkan("tekil")
gunc=efkan("cogul")
End If

topt=efkan("tekil")+topt
topc=efkan("cogul")+topc
efkan.movenext
loop
efkan.close 

gunluktekilortalama = topt / toplamgun
gunlukcogulortalama = topc / toplamgun

Session("ziyaret")="yes"
%>   



<!-- TABLO BAÞLIYOR -->
<table  width="100%"  bordercolor="#CCFFFF" border="0" cellspacing="0" cellpadding="0">
<tr height="20">
<td background="forumimg/mn.gif" width="100%" align="center" valign="center" >
<FONT COLOR="#FFFFFF"><B>Üye Ýstatistikleri</B></FONT>
</td></tr>

<!-- AKTÝF ÜYELER -->
<tr bgcolor="<%=bgcolor1%>" height="20"><td width="100%" align="left" valign="center">
<IMG SRC="images/onn.gif" WIDTH="11" HEIGHT="11" BORDER="0" ALT="">
<A HREF="default.asp?part=uyegorev&gorev=uyeler&diz=sontarih"><B>Online Üyeler</B></A></td></tr>
<tr bgcolor="<%=bgcolor2%>" height="40"><td width="100%" align="left" valign="top">
<% 'ONLÝNE ÜYELER
Session.LCID = 1055
DefaultLCID = Session.LCID 
onlineuye=0
sor="SELECT * FROM uyeler where onay=1 "
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("n",efkan("sontarih"),now)  ' ÞU AN DAN 1 DAKKA CIKAR SON TARÝH FARKI BÜYÜKSE
if zaman > 1 then
'boþ
else 
onlineuye=onlineuye + 1
%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan("id")%>">
<FONT COLOR="#3300CC"><%=efkan("kadi")%></FONT></A>,
<% 
End If
efkan.movenext 
loop 
efkan.close
If onlineuye=0 Then Response.Write "Online Üyemiz Yok" End If
%>
</td></tr>




<!-- SON 24 SAATTE AKTÝF OLANLAR -->
<tr bgcolor="<%=bgcolor1%>" height="20"><td width="100%" align="left" valign="center">
&nbsp;<IMG SRC="images/pin.gif" WIDTH="13" HEIGHT="15" BORDER="0" ALT="">
<A HREF="default.asp?part=uyegorev&gorev=uyeler&diz=sontarih"><B>Bugün kimler vardý</B></A></td></tr>
<tr bgcolor="<%=bgcolor2%>" height="40"><td width="100%" align="left" valign="top">
<% Session.LCID = 1055
DefaultLCID = Session.LCID 
sor="SELECT * FROM uyeler where onay=1 order by sontarih desc "
efkan.Open sor,Sur,1,3
do while not efkan.eof  
zaman=datediff("h",efkan("sontarih"),now)  ' ÞU ANLA KAYITLI SONTARÝH FARKI 24 DEN BÜYÜKSE
if zaman > 24 then
'BOÞ BIRAK
else %>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan("id")%>">
<FONT COLOR="#3300CC"><%=efkan("kadi")%></FONT></A>,
<% 
End If
efkan.movenext 
loop 
efkan.close
%>
</td></tr>



<!-- EN AKTÝFLER -->
<tr bgcolor="<%=bgcolor1%>" height="20"><td width="100%" align="left" valign="center">
&nbsp;<IMG SRC="images/pin.gif" WIDTH="13" HEIGHT="15" BORDER="0" ALT="">
<A HREF="default.asp?part=uyegorev&gorev=uyeler&diz=hit"><B>En Aktif Üyeler</B></A></td></tr>
<tr bgcolor="<%=bgcolor2%>" height="40"><td width="100%" align="left" valign="top">
<% sor="SELECT * FROM uyeler where onay=1 and hit <> 0 order by hit desc "
efkan.Open sor,Sur,1,3
for i=1 to 30
if efkan.eof then exit for%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan("id")%>">
<FONT COLOR="#3300CC"><%=efkan("kadi")%></FONT></A>,
<% 
efkan.movenext 
next 
efkan.close
%>
</td></tr>



<!-- EN SON ÜYELER -->
<tr bgcolor="<%=bgcolor1%>" height="20"><td width="100%" align="left" valign="center">
&nbsp;<IMG SRC="images/pin.gif" WIDTH="13" HEIGHT="15" BORDER="0" ALT="">
<A HREF="default.asp?part=uyegorev&gorev=uyeler&diz=id"><B>En Yeni Üyeler</B></A></td></tr>
<tr bgcolor="<%=bgcolor2%>" height="50"><td width="100%" align="left" valign="top">
<%sor="SELECT * FROM uyeler where onay=1 order by id desc "
efkan.Open sor,Sur,1,3
for i=1 to 35
if efkan.eof then exit for%>
<A HREF="default.asp?part=uyegorev&gorev=uyebilgi&id=<%=efkan("id")%>">
<FONT COLOR="#3300CC"><%=efkan("kadi")%></FONT></A>,
<% 
efkan.movenext 
next
efkan.close
%></td></tr>


<tr bgcolor="<%=bgcolor1%>" height="40"><td width="100%" align="center" valign="center">


<!-- AKTÝF ADET ZÝYARETCÝ-->
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Aktif Ziyaretçi  <B><%=onlineadet%></B>

<!--AKTÝF  ÜYE ADET -->
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Aktif Üye <B><%=onlineuye%></B>

<!-- ÜYE SAYIMIZ -->
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
<A HREF="default.asp?part=uyegorev&gorev=uyeler">Üye Sayýmýz</A> <B><%=toplamuye%></B>
<BR>
<!-- TEKÝL ÇOÐUL SAYAC BAÞLA -->
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Bugün Tekil <B><%=gunt%></B>
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Bugün Çoðul <B><%=gunc%></B>
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Toplam Tekil <B><%=topt%></B>
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Toplam Çoðul <B><%=topc%></B>
<IMG SRC="images/tekil.gif" WIDTH="17" HEIGHT="14" BORDER="0" ALT="">
Ip No: <%=ip_no%>
</td></tr>

</table>




<%
Set efkan1=Nothing
Set efkan=Nothing
%>

