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
<% 
Response.Buffer = True 
id            =kontrol(Request.QueryString("id"))
guvenlik    =kontrol(Request.QueryString("guvenlik"))
grv          =temizle(Request.QueryString("grv"))
isl           =temizle(Request.QueryString("isl"))

'EMAÝLLE SORU SÝLME ONAYLAMA
If isl="soru" Then 
sor = "select  * from  sorular where id="&id&" and guvenlik="&guvenlik&"   "
forum.Open sor,forumbag,1,3

if forum.eof or forum.bof then
Response.Write "<B>Kayýt Bulunamadý....</B>"
Response.End
End If

If grv="onay" then
forum("onay")=1
forum.Update
forum.close
Else
forum.close
sor = "DELETE from sorular where id="&id&" and guvenlik="&guvenlik&"   "
forum.Open sor,forumbag,1,3
End If
Response.Write "<B>" & isl & " &nbsp; " & grv & "<B>  iþleminiz baþarý ile yapýldý....</B>"
End If




'EMAÝLLE CEVAP SÝLME ONAYLAMA
If isl="cevap" Then 
sor = "select  * from  cevaplar where id="&id&" and guvenlik="&guvenlik&"   "
forum.Open sor,forumbag,1,3

if forum.eof or forum.bof then
Response.Write "<B>Kayýt Bulunamadý....</B>"
Response.End
End If

If grv="onay" then
forum("onay")=1
forum.Update
forum.close
Else
forum.close
sor = "DELETE from cevaplar where id="&id&" and guvenlik="&guvenlik&"   "
forum.Open sor,forumbag,1,3
End If
Response.Write "<B>" & isl & " &nbsp; " & grv & "<B>  iþleminiz baþarý ile yapýldý....</B>"
End If


%>