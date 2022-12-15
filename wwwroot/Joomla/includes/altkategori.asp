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
'		Tel : 0544 275 9804 - 0537 275 3655
'		Mail: emre06@hotmail.com.tr / info@joomlasp.com/.net/.org
'
'
'		Lisans Anlaþmasý Gereði Lütfen Google Reklam Bölümünü Sitenizden kaldýrmayýnýz. Bu sizin GOOGLE reklamlarýný yapmanýza
'		kesinlikle bir engel deðildir. reklam.asp içeriðinin yada yayýnladýðý verinin deðiþmesi lisans politikasýnýn dýþýna çýkýlmasýna
'		ve JoomlASP CMS sistemini ücretsiz yayýnlamak yerine ücretlie hale getirmeye bizi teþfik etmektedir. Bu Sistem için verilen emeðe
'		saygý ve bir çeþit ödeme seçeneði olarak GOOGLE reklamýmýzýn deðiþtirmemesi yada silinmemesi gerekmektedir.
%>
<table width="100%" border="0" cellpadding="0" cellspacing="10"><tr align="left">
<%
katid = Request.QueryString ("katid")
IF Not IsNumeric(Request.QueryString ("katid")) THEN
response.Redirect "hata.asp"
End if

deste = vsayi
If Request.QueryString("s") <> "" Then 
  Sayfa = CInt(Request.QueryString("s"))
Else 
  Sayfa = 1 
End If 
set veri = baglanti.Execute("select * from gop_veriler where katid=" & guvenlik(request.querystring("katid")) & " order by vid DESC LIMIT "& (deste*Sayfa)-(deste) & "," & deste)

Set SQLToplam = baglanti.Execute("select count(vid) from gop_veriler where katid=" & guvenlik(request.querystring("katid"))&"") 
TopKayit = SQLToplam(0) 
if veri.eof or veri.bof then
Response.Write not_data
else

for z=1 to deste
if  veri.eof then exit for



Response.Write "<td width=""50%"" valign=""top""><b><img src=upload/"& veri("vresim") &" height=""75"" width=""100"" align=""left"" onerror=""this.src='images/joomlasp.jpg'"" align=left height=75 width=100>" & veri("vbaslik") & "...</b><br>"&left(veri("vicerik"),200)&"... <a href=default.asp?islem=oku&vid="&veri("vid")&">"&read_more&"</a><br>"&hit&": "&veri("vhit")&"</td>"

if z mod 2 = 0 then response.write "</tr><tr>"


veri.MoveNext
next
veri.close
set veri=nothing
end if
%>
</tr></table>
<div align="center">
<%
If CInt(TopKayit) > CInt(deste) Then 
SayfaSayisi = CInt(TopKayit) / CInt(deste) 
If InStr(1,SayfaSayisi,",",1) > 1 Then SayfaSayisi = CInt(Left(SayfaSayisi,InStr(1,SayfaSayisi,",",1))) + 1 

If SayfaSayisi > 1 Then 
Response.Write "<a class=pagenav href=default.asp?islem=altkategori&katid="& guvenlik(request.querystring("katid"))&"&s=1><< "&start_page&"</a> "
if Sayfa = 1 then
Response.Write ""
else
Response.Write "<a class=pagenav href=default.asp?islem=altkategori&katid="& guvenlik(request.querystring("katid"))&"&s="&Sayfa-1&">< "&prev_page&"</a> "
end if

for t=d to Sayfa-1
if Sayfa > Sayfa-5 and t > Sayfa-5  and t > 0 then
Response.Write " <a class=pagenav href=default.asp?islem=altkategori&katid="& guvenlik(request.querystring("katid"))&"&s=" & t & "><b>" & t & "</b></a> "
end if
next 

For d=Sayfa To Sayfa+4
if d = CInt(Sayfa) then
Response.Write "<span class=pagenav><b>" & d & "</b></span>"
elseif d > SayfaSayisi then
Response.Write ""
else
Response.Write " <a class=pagenav href=default.asp?islem=altkategori&katid="& guvenlik(request.querystring("katid"))&"&s=" & d & "><b>" & d & "</b></a> " 
end if
Next


if Sayfa = SayfaSayisi then
Response.Write ""
else
Response.Write " <a class=pagenav href=default.asp?islem=altkategori&katid="& guvenlik(request.querystring("katid"))&"&s="&Sayfa+1&"> "&next_page&" ></a> "
Response.Write " <a class=pagenav href=default.asp?islem=altkategori&katid="& guvenlik(request.querystring("katid"))&"&s="&SayfaSayisi&">"&end_page&" >></a>"
end if


End If 
End If 
%>
</div>