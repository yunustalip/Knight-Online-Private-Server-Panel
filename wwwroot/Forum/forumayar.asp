<LINK href="ie.css" type=text/css rel=stylesheet>
<% 
'DB YOLLARINI BURADAN AYARLAYIN Y�NET�M B�L�M�NDEK� YOLLARI  BA� DOSYALARINDAN AYARLAYIN
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

hemenyayinla      =1        'E�er admin onay� istenirse 0 hemen yay�nlamak i�in 1
emaildogrulama    =0                 'AKT�VASYONLU �YEL�K A�IK ���N 1  D�REK �YE OLMA ���N 0
uyeresimyolu       = "/uyeler/"      '�YE RES�M UPLOAD KLAS�R�  YAZMA YETK�L� KLAS�R OLMALI
fotoyukleme        =1                '�YELER RES�M YUKLEME 1 A�IK 0 KAPALI
uyeresimadet      =1                 'UYE MAX�MUM KA� RES�M UPLOAD EDEB�L�R
sayfaveri           =20           ' B�R SAYFADAK� VER� ADED� BAZI YERLERDE KULLANMAMI� OLABL�R�M
hitbirim              =20     '�YELER�N H�T� YILDIZ KATSAYISI �RNEK 50 G�R�� ���N B�R YILDIZ KOY G�B�

websayfam      ="http://www.siteniz.com"  'BURAYA GER�EK URLN�Z� YAZIN FORUM NERDEYSE O
emailadresim    ="email@email.com"   'EMA�L ADRES�N�Z 
emailadresim1  = emailadresim   'AYNI KALSIN
mailhost          ="localhost"     '  AYNI KALSIN
'mailhost          ="127.0.0.1"
ipm                 =""  'HOST IP NO  GEREK YOK 
title               ="Forum sayfam�za ho�geldiniz.....Efkan Forum v.4.3 "
keywords         ="cnc,cad,cam,bilgisayar,donan�m,e�itim,konular,spor,nokia,siemens,internet,siteler,sanat,sinema,tiyatro,politika,i� d�nyas�,siyaset,asp,java,html,php,matematik,fizik,borsa,finans,liderler,�lkeler,b�lgeler,her �ey"

'TEMALAR  RENKLER� DE���T�REB�L�RS�N�Z KEND�N�ZE G�RE
If session("tema")="1" then
bgcolor1 = "#EAEAEA"  'L�STELEME RENG� 
bgcolor2 = "#F9F9F9"            'L�STELEME RENG� 
ElseIf session("tema")="2" Or session("tema")="" then
bgcolor1 = "#E2E2E2"  'L�STELEME RENG� 
bgcolor2 = "#F3F3F3"            'L�STELEME RENG� 
ElseIf session("tema")="3" then
bgcolor1 = "#66FFCC"  'L�STELEME RENG� 
bgcolor2 = "#FFFFCC"            'L�STELEME RENG� 
ElseIf session("tema")="4" then
bgcolor1 = "#FFFFFF"  'L�STELEME RENG� 
bgcolor2 = "#FFFFFF"   
ElseIf session("tema")="5" then
bgcolor1 = "#66CCFF"  'L�STELEME RENG� 
bgcolor2 = "#F3F3F3"   
End If

sub emailgonder(email,konu,emesaj)

On Error Resume Next 

SELECT Case 2   'BURAYA A�A�IDAK� KOMPANENTLERE G�RE DE����TR�N   �RNEK CDONTS ���N 1 YAZIN
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

Case 4  'ASP MA�L
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