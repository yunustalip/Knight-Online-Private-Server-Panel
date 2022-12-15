<LINK href="ie.css" type=text/css rel=stylesheet>
<% 
'DB YOLLARINI BURADAN AYARLAYIN YÖNETÝM BÖLÜMÜNDEKÝ YOLLARI  BAÐ DOSYALARINDAN AYARLAYIN
chatyolu="../dbb/chat.mdb"
Set forumbag = Server.CreateObject("ADODB.Connection")
forumbag.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../dbb/forum.mdb")
Set forum= Server.CreateObject("ADODB.Recordset")
Set forum1= Server.CreateObject("ADODB.Recordset")
Set forum2= Server.CreateObject("ADODB.Recordset")
Set forum3= Server.CreateObject("ADODB.Recordset")

Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../dbb/uyeler.mdb")
Set efkan = Server.CreateObject("ADODB.Recordset")
Set efkan1 = Server.CreateObject("ADODB.Recordset")
Set efkan2 = Server.CreateObject("ADODB.Recordset")

hemenyayinla      =1        'EÐer admin onayý istenirse 0 hemen yayýnlamak için 1
emaildogrulama    =0                 'AKTÝVASYONLU ÜYELÝK AÇIK ÝÇÝN 1  DÝREK ÜYE OLMA ÝÇÝN 0
uyeresimyolu       = "/uyeler/"      'ÜYE RESÝM UPLOAD KLASÖRÜ  YAZMA YETKÝLÝ KLASÖR OLMALI
fotoyukleme        =1                'ÜYELER RESÝM YUKLEME 1 AÇIK 0 KAPALI
uyeresimadet      =1                 'UYE MAXÝMUM KAÇ RESÝM UPLOAD EDEBÝLÝR
sayfaveri           =20           ' BÝR SAYFADAKÝ VERÝ ADEDÝ BAZI YERLERDE KULLANMAMIÞ OLABLÝRÝM
hitbirim              =20     'ÜYELERÝN HÝTÝ YILDIZ KATSAYISI ÖRNEK 50 GÝRÝÞ ÝÇÝN BÝR YILDIZ KOY GÝBÝ

websayfam      ="http://www.siteniz.com"  'BURAYA GERÇEK URLNÝZÝ YAZIN FORUM NERDEYSE O
emailadresim    ="email@email.com"   'EMAÝL ADRESÝNÝZ 
emailadresim1  = emailadresim   'AYNI KALSIN
mailhost          ="localhost"     '  AYNI KALSIN
'mailhost          ="127.0.0.1"
ipm                 =""  'HOST IP NO  GEREK YOK 
title               ="Forum sayfamýza hoþgeldiniz.....Efkan Forum v.4.3 "
keywords         ="cnc,cad,cam,bilgisayar,donaným,eðitim,konular,spor,nokia,siemens,internet,siteler,sanat,sinema,tiyatro,politika,iþ dünyasý,siyaset,asp,java,html,php,matematik,fizik,borsa,finans,liderler,ülkeler,bölgeler,her þey"

'TEMALAR  RENKLERÝ DEÐÝÞTÝREBÝLÝRSÝNÝZ KENDÝNÝZE GÖRE
If session("tema")="1" then
bgcolor1 = "#EAEAEA"  'LÝSTELEME RENGÝ 
bgcolor2 = "#F9F9F9"            'LÝSTELEME RENGÝ 
ElseIf session("tema")="2" Or session("tema")="" then
bgcolor1 = "#E2E2E2"  'LÝSTELEME RENGÝ 
bgcolor2 = "#F3F3F3"            'LÝSTELEME RENGÝ 
ElseIf session("tema")="3" then
bgcolor1 = "#66FFCC"  'LÝSTELEME RENGÝ 
bgcolor2 = "#FFFFCC"            'LÝSTELEME RENGÝ 
ElseIf session("tema")="4" then
bgcolor1 = "#FFFFFF"  'LÝSTELEME RENGÝ 
bgcolor2 = "#FFFFFF"   
ElseIf session("tema")="5" then
bgcolor1 = "#66CCFF"  'LÝSTELEME RENGÝ 
bgcolor2 = "#F3F3F3"   
End If

sub emailgonder(email,konu,emesaj)

On Error Resume Next 

SELECT Case 2   'BURAYA AÞAÐIDAKÝ KOMPANENTLERE GÖRE DEÐÝÞÝTRÝN   ÖRNEK CDONTS ÝÇÝN 1 YAZIN
Case 1 
set mail = Server.CreateObject("CDONTS.Newmail")
mail.To = "<" &email& ">"
'mail.cc = emailadresim1
mail.cc = "<" & emailadresim1 & ">"
mail.From ="<" & emailadresim & ">"
'mail.From =websayfam & "<" & emailadresim & ">"
mail.Subject = konu
mail.BodyFormat = 0  '0 html   1 text
mail.mailFormat = 0  '0 html   1 text
mail.Body = emesaj
mail.Importance = 2 '0=Low, 1=Normal, 2=High
mail.Send  
Set mail = Nothing

Case 2
Set mail = Server.CreateObject("Persits.MailSender")
mail.IsHTML = True  'html
mail.Host = mailhost
mail.From = emailadresim
'mail.FromName =websayfam
mail.AddAddress email
mail.Subject = konu
mail.Body = emesaj
mail.Send
Set mail = Nothing

Case 3
Set mail = Server.CreateObject("Jmail.Message")
mail.AddRecipient email , emailadresim
mail.From = emailadresim
mail.Subject = konu
mail.HTMLBody = "<html><body>"& emesaj &"</body></html>"
mail.Send( "LOCALHOST" )
Set mail = Nothing

Case 4  'ASP MAÝL
Set mail = Server.CreateObject("SMTPsvg.mailer")
mail.FromAddress = emailadresim
mail.FromName = websayfam
mail.RemoteHost =mailhost
mail.AddRecipient = email
mail.Subject = konu
mail.BodyText = emesaj
mail.Sendmail
Set mail = Nothing

Case 5  'CDO MESAJ
Const cdoSendUsingPort = 2 
StrSmartHost = ipm
set iMsg = CreateObject("CDO.Message") 
set iConf = CreateObject("CDO.Configuration") 
Set Flds = iConf.Fields 
' set the CDOSYS configuration fields to use port 25 on the SMTP server 
With Flds 
.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort 
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmartHost 
.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10 
.Update 
End With 
With iMsg 
Set .Configuration = iConf 
.BodyFormat = 0  '0 html   1 text 
.MailFormat = 0  '0 html   1 text 
.To = email
.cc = emailadresim1
.From = websayfam
.Subject = konu
.host =mailhost
.Importance = 1 '0=Low, 1=Normal, 2=High 
.htmlBody =emesaj
.Send 
End With 

Case 6  'bamboo
set mail = Server.CreateObject("Bamboo.SMTP")
mail.Server = mailhost
mail.Rcpt = email
mail.From = emailadresim
mail.FromName = websayfam
mail.Subject = konu
mail.Message =emesaj
mail.Send
set mail = Nothing
End Select

end sub

%>




<script language=JavaScript>
function SayiKontrol(ths) 
{ 
        if (event.keyCode < 46 || event.keyCode > 57) 
    { 
        event.keyCode = 0; 
        return  false; 
    } 
    else 
        return true; 
} 
</script>

<!--#INCLUDE file="fonksiyon.asp"-->