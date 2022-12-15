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
<!--#include file="../functions/fonksiyonlar.asp"-->
<!--#include file="md5.asp"-->

<%
uye_mail = request("uye_mail")
if uye_mail= "" then
%>
<style type="text/css">
<!--
.style3 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12; }
.style4 {font-size: 12}
-->
</style>

<form action="" method="get">
<table width="100%" border="0" cellpadding="2" cellspacing="2">
  <tr>
    <td colspan="3"><div align="center" class="style3" style="font-weight: bold"><%= lost_pass %></div></td>
  </tr>
  <tr>
    <td width="43%"><span class="style3"><%= email %> </span></td>
    <td width="1%"><span class="style3">:</span></td>
    <td width="56%"><input name="uye_mail" type="text" id="uye_mail" size="30" /></td>
  </tr>
  <tr>
    <td colspan="3"><span class="style3">
      <%
response.Write "<div align=center>"&notice8&" </div>"
%>
    </span></td>
    </tr>
  <tr>
    <td><span class="style4"></span></td>
    <td><span class="style4"></span></td>
    <td><input name="Submit" type="submit" value="<%= sent_pass %>" /></td>
  </tr>
</table>

</form>

<%
else
Function SifreUret(Uzunluk)
Karakterler = "0123456789abcdefghijklmnoprqstuvyzABCDEFGHIJKLMNOPRQSTUVYZ"
Randomize
KarakterBoyu = Len(Karakterler)
For i = 1 To Uzunluk
      KacinciKarakter = Int((KarakterBoyu * Rnd) + 1)
      UretilenSifre = UretilenSifre & Mid(Karakterler,KacinciKarakter,1)
Next
SifreUret = UretilenSifre
End Function
sifreniz = SifreUret(6)
%>
<%
'Mail adresinden gelen bilgiyi iþle
dim oku
Set oku = baglanti.Execute("Select * from gop_uyeler where uye_mail='" & uye_mail & "' ;")
if oku.eof or oku.bof then
Response.Redirect "404.asp"
else
baglanti.Execute("UPDATE gop_uyeler set uye_sifre='"&md5(sifreniz)&"' where uye_mail='" & uye_mail & "' ;")
end if

'mail adresinden gelen veri iþlendi
%>

<%
' Email
Set objCDOSYSMail = Server.CreateObject("CDO.Message")
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
' SMTP Ayarlarý
HTML = "<hr><a href=""http://www.joomlasp.com"" target=_blank>JoomlASP</a> | Geliþime Açýk site Yönetimi Sistemi " &surum&"<br><br>"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objCDOSYSCon.Fields.Update
' CDOSYS Ayarlarý
Set objCDOSYSMail.Configuration = objCDOSYSCon

objCDOSYSMail.From = site_mail
objCDOSYSMail.To = uye_mail
objCDOSYSMail.Subject = lost_pass
objCDOSYSMail.HTMLBody = "Sayýn; " &oku("uye_adi")& "<br>Þifreniz: "&sifreniz&" <br><br>Ýsteðiniz Üzerine Þifreniz Yukarýdaki gibi düzenlenmiþtir.<br>"&HTML

' Gönder
objCDOSYSMail.Send

' Herþeyi Kapat
Set objCDOSYSMail = Nothing
Set objCDOSYSCon = Nothing
oku.close 
set oku=nothing 
If err Then ' hata mesajýný alalým Mail Gönderilmemiþse..
Response.Write err.Description & "<br>" & not_sent_message
Else ' Mail Gönderilmiþ ise
Response.Write "<br><br><center><b><font color=green>"&successful&"</font></b><br><font color=red>"&notice9&"</center></font><br><br>"
End If
end if

%> 