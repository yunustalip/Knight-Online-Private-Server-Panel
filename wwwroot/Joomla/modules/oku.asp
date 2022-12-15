<%
'      JoomlASP Site Yönetimi Sistemi (CMS)
'
'      Copyright (C) 2007 Hasan Emre ASKER
'
'      This program is free software; you can redistribute it and/or modify it
'      under the terms of the GNU General Public License as published by the Free
'      Software Foundation; either version 3 of the License, or (at your option)
'      any later version.
'
'      This program is distributed in the hope that it will be useful, but WITHOUT
'      ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
'      FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
'      You should have received a copy of the GNU General Public License along with
'      this library; if not, write to the JoomlASP Asp Yazýlým Sistemleri., Kargaz Doðal Gaz Bilgi Ýþlem Müdürlüðü
'       36100 Kars / Merkez 
'		Tel : 0544 275 9804
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.

vid = Request.QueryString ("vid")
IF Not IsNumeric(Request.QueryString ("vid")) THEN
response.Redirect "hata.asp"

End if

Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""contentpaneopen"">"
dim oku
Set oku = baglanti.Execute("Select * from gop_veriler where vid=" & guvenlik(request.querystring("vid")) & " ;")
if oku.eof or oku.bof then
Response.Redirect "404.asp"
else
Response.Write "<title>"& oku("vbaslik")&"</title>"
vhit = oku("vhit")
baglanti.Execute("UPDATE gop_veriler set vhit='"&vhit+1&"' where vid=" & guvenlik(request.querystring("vid")) & " ;")
Response.Write "<tr><td><b>"& oku("vbaslik") &"</b></td></tr><tr><td>"&oku("vicerik")&"<br>"

SQLyazar ="SELECT * FROM gop_uyeler where uye_id ='"& guvenlik(oku("uye_id")) &"';"
set yazar = server.createobject("ADODB.Recordset")
yazar.open SQLyazar , Baglanti
Response.Write "<b>"&written_by&":</b>"&uyeisimkontrol(yazar("uye_adi"))&"<br><b>"&date_in&":</b> "&oku("vtarih")&"<br><b>"&hit&":</b> "&oku("vhit")&"<br></td></tr>"

yazar.close
set yazar=nothing

Response.Write "</table><br><b>"&tags&" : </b>"


dim bolunecek
kelimelerimiz = oku("vetiket")
if kelimelerimiz = "" then
Response.Write "Yok"
else

bolunecek =  Split(kelimelerimiz, ",")
    for  bb = 0 to  Ubound(bolunecek)
        strEKRAN = strEKRAN  & "<a href=""?islem=etiket&tag="&Trim(bolunecek(bb ))&""">"& bolunecek(bb ) &"</a>" & ", "
    next
Response.Write Left(strEKRAN,(Len(strEKRAN)-2))

end if

response.write "<br><br>"
%>


<script language="Javascript" type="text/javascript">
var strTitle = encodeURIComponent("<%=oku("vbaslik") %>");
var strTitle2 = "<%=oku("vbaslik") %>";
var strURL = encodeURIComponent(location.href);
var strThumb = "<%=oku("vresim") %>";
var strDesc = encodeURIComponent("");
var strDesc2 = encodeURIComponent("");
var strKeyword = encodeURIComponent("<%=oku("vetiket") %>");
var strKeyword2 = encodeURIComponent("<%=oku("vetiket") %>");
</script>

<div id="altnav1">
<ul>
<li style="padding: 10px 0 5px 2px;">Paylaþ : <a href="javascript:void(0);" onclick="addDelicious(strTitle, strURL, strDesc, strKeyword);"><img src="icons/delicious.png" alt="del.icio.us'a ekle" /></a>&nbsp;<a href="void(0);" onclick="addGBookmark(strTitle, strURL, strDesc, strKeyword2);"><img src="icons/google.png" alt="Google yerimlerine ekle" /></a>&nbsp;<a href="javascript:void(0);"  onclick="addDigg(strTitle, strURL, strDesc);"><img src="icons/digg.png" alt="Digg.it'e ekle" /></a>&nbsp;<a href="javascript:void(0);" onclick="addYahoo(strTitle2, strURL, strDesc, strKeyword2);"><img src="icons/yahoo_myweb.png" alt="Yahoo! MyWeb'e ekle" /></a>&nbsp;<a href="javascript:void(0);"  onclick="addTusul(strURL);"><img src="icons/tusuldat.png" alt="Tusuldat!" /></a>&nbsp;<a href="javascript:void(0);"  onclick="addFacebook(strThumb,strURL,strTitle,strDesc);"><img src="icons/facebook.png" alt="Facebook'ta gönder" /></a>&nbsp;<a href="javascript:void(0);"  onclick="addTechnorati(strURL);"><img src="icons/technorati.png" alt="Technorati'ye ekle" /></a>&nbsp;<a href="javascript:void(0);"  onclick="addSpurl(strTitle, strURL);"><img src="icons/spurl.png" alt="Spurl'a ekle" /></a></li></ul></div><div id="altnav2" style="height: 93px;"><div style="width: 200px; padding-top: 15px; vertical-align: middle;" align="center"></div></div><br style="clear: both;" />

<%
end if
oku.close 
set oku=nothing


if "goster" = vyorum then
Response.Write "<hr size=1><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td bgcolor=""#efefef""><center><strong>"&comments&"</strong></center></td></tr></table><hr size=1>"

if Session("durum")="giris_yapmis" then
uye_adi=session("uye_adi")
mesaj_no = request.querystring("vid")
set uye_isim = baglanti.Execute("select * from gop_uyeler where uye_adi = '" & uye_adi & "';")

Response.Write "<form name=""form1"" method=""post"" action=""default.asp?islem=yorum_gonder&vid="&mesaj_no&"""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td width=""18%"">"&username&"</td><td width=""2%"">:</td><td width=""80%""><strong>"& uyeisimkontrol(uye_isim("uye_adi"))&"</strong><br></td></tr><tr><td>"&comment&"</td><td>:</td><td><textarea class=inputbox name=""yorum"" id=""yorum"" cols=""45"" rows=""10""></textarea></td></tr><tr><td>&nbsp;</td><td colspan=""2""><input class=button type=""submit"" name=""button"" id=""button"" value="""&add_comment&"""></td></tr></table></form>"
else
Response.Write ""
end if

deste = 15
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 
set veri = baglanti.Execute("select * from gop_yorumlar where vid ='"& guvenlik(request.querystring("vid")) &"' and yorum_onay = '"& 1 &"' order by yorum_id DESC LIMIT "& (deste*Sayfa)-(deste) & "," & deste)

if veri.eof or veri.bof then
Response.Write "<br><br><br><center><b>"&not_comments&"</b></center><br><br>"
else

Set SQLToplam = baglanti.Execute("select count(yorum_id) from gop_yorumlar where vid ='"& guvenlik(request.querystring("vid")) &"'") 
TopKayit = SQLToplam(0) 

for z=1 to deste
if  veri.eof then exit for

set ekleyen = baglanti.Execute("select * from gop_uyeler where uye_id ='"& guvenlik(veri("uye_id")) &"';")

if ekleyen.eof or ekleyen.bof then
Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td width=""60%""><strong>"&added&" : "&not_member&" </strong></td><td width=""40%""><strong>"&date_in&" :"&veri("yorum_tarih")&"</strong></td></tr><tr><td colspan=""2"">"&guvenlikyorum(veri("yorum"))&"</td></tr></table><br><hr size=1 /><br>"

else

Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td width=""60%""><strong>"&added&" :"&uyeisimkontrol(ekleyen("uye_adi"))&"</strong></td><td width=""40%""><strong>"&date_in&" :"&veri("yorum_tarih")&"</strong></td></tr><tr><td colspan=""2"">"&guvenlik(veri("yorum"))&"</td></tr></table>"
if z mod 1 = 0 then response.write "<br><hr size=1 /><br>"
end if


veri.MoveNext
next
veri.close
set veri=nothing
end if


else
end if
%>