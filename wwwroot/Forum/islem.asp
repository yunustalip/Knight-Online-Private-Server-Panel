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
<% 
Response.Buffer = True 
id            =kontrol(Request.QueryString("id"))
guvenlik    =kontrol(Request.QueryString("guvenlik"))
grv          =temizle(Request.QueryString("grv"))
isl           =temizle(Request.QueryString("isl"))

'EMA�LLE SORU S�LME ONAYLAMA
If isl="soru" Then 
sor = "select  * from  sorular where id="&id&" and guvenlik="&guvenlik&"   "
forum.Open sor,forumbag,1,3

if forum.eof or forum.bof then
Response.Write "<B>Kay�t Bulunamad�....</B>"
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
Response.Write "<B>" & isl & " &nbsp; " & grv & "<B>  i�leminiz ba�ar� ile yap�ld�....</B>"
End If




'EMA�LLE CEVAP S�LME ONAYLAMA
If isl="cevap" Then 
sor = "select  * from  cevaplar where id="&id&" and guvenlik="&guvenlik&"   "
forum.Open sor,forumbag,1,3

if forum.eof or forum.bof then
Response.Write "<B>Kay�t Bulunamad�....</B>"
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
Response.Write "<B>" & isl & " &nbsp; " & grv & "<B>  i�leminiz ba�ar� ile yap�ld�....</B>"
End If


%>