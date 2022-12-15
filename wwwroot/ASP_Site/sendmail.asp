<% @ Language="VBScript" %>
<% Option Explicit %>
<% 
' ASP ActiveX Mail Bileseni Nesnesi (CDONTS, Persits, ASPEmail)      
' 16.01.2005  -  Sunday    
' deathrole[at">msn[dot">com    
' ASP Rehberi - www.asprehberi.net                

    'Hata olursa bir sonraki satirdan devam et.
    On Error Resume Next

    Dim COMType              'Kullanilacak olam mail bileseni. KULLANACAGINIZ BILESENIN BASINDAKI "'" KALDIRIN !!!
    Dim objMail              'Mail göndermemizi saglayan sunucu nesnesi
    Dim blnHTMLMail          'Mailin HTML/Text formati
    Dim strBody              'Gönderilecek mesaj
    Dim YourName            'Nesneye ait gönderenin adi
    Dim FromEmail            'Nesneye ait gönderen mail adresi
    Dim ToEmail              'Nesneye ait giden mail adresi
    Dim MailServer          'Nesneye ait mail sunucusu
    Dim MailSubject          'Nesneye ait mail konusu
    Dim MailBody            'Nesneye ait mail mesaji

    'Degiskenlere degerlerini veriyoruz
    'Bilesen seçimini yapin
    COMType = "CDONTS"
    'COMType = "Persits"
    'COMType = "ASPEmail"
    blnHTMLMail = True
    strBody = "<html><h1>Deneme Yazýsý</h1></html>"
    YourName = "Deathrole"
    FromEmail = "user"
    ToEmail = "mail@domain.com"
    MailServer = "mail.domain.com"
    MailSubject = "Mail Konusu"
    MailBody = strBody


    'Mail bilesenimize göre nesnemizi olusturmamiza yardimci olan select ifadesi ile sinama islemi yapiyoruz.
    Select Case COMType

    'Eger bilesen CDONTS ise,
    Case "CDONTS"

        'Nesnemizi olusturalim
Set objMail = Server.CreateObject("CDONTS.NewMail")

'Nesnemizin özelliklerini belirliyoruz.
With objMail
      If blnHTMLMail then
      .MailFormat = HTML
      .BodyFormat = HTML
      Else
      .MailFormat = Text
      .BodyFormat = Text
      End If
      .From = YourName & " <" & FromEmail & ">"
      .To = ToEmail
      .Subject = MailSubject
      .Body = MailBody
      .Send
End With


    'Eger bilesen Persits ise,
    Case "Persits"

        'Nesnemizi olusturalim
Set objMail = Server.CreateObject("Persits.MailSender")

'Nesnemizin özelliklerini belirliyoruz.
With objMail
      If blnHTMLMail then
      .IsHTML = True
      Else
      .IsHTML = False
      End If
      .From = FromEmail
      .FromName = YourName
      .Host = MailServer
      .Subject = MailSubject
      .Body = MailBody
      .Send
End With


    'Eger bilesen ASPMail ise,
    Case "ASPMail"

        'Nesnemizi olusturalim
Set objMail = Server.CreateObject("SMTPsvg.Mailer")

'Nesnemizin özelliklerini belirliyoruz.
With objMail
      .FromAddress = FromEmail
      .FromName = YourName
      .RemoteHost = MailServer
      .AddRecipient = ToEmail
      .Subject = MailSubject
      .BodyText = MailBody
      .SendMail
End With
    End Select

    'Nesnemizi kaldiriyoruz.
Set objMail = Nothing
%> 

