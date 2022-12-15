<!--#include file="_inc/DbSettings.Asp"-->
<!--#include file="_inc/Connect.Asp"-->
<%Response.Charset = "iso-8859-9"
Dim objfso, txtfile, txtToDisplay

Set objfso = Server.CreateObject("Scripting.FileSystemObject")
activeusers=split(APPLICATION("AktifKullaniciListesi"),"|")

If IsArray(Application("DeathLog")) Then
DeathLog=Application("DeathLog")
Else
ReDim DeathLog(11,1)
DeathLog(0,0)=Now()-1
DeathLog(1,0)=0
End If

If DateDiff("s",DeathLog(0,0),Now())>15 Then

If objfso.FileExists("D:\KO\SERVER FILES\3 - Ebenezer\DeathLog-"&year(now)&"-"&month(now)&"-"&day(now)&".txt")=True Then

Set txtfile = objfso.GetFile("D:\KO\SERVER FILES\3 - Ebenezer\DeathLog-"&year(now)&"-"&month(now)&"-"&day(now)&".txt")

If DeathLog(1,0)<>txtfile.size Then

Set Ag = txtfile.OpenAsTextStream(1,-2)

If Not ag.AtEndOfStream Then
Satir=Split(Ag.ReadAll,vbCrlf)
If Session("DeathLog")="" Then
Session("DeathLog")=UBound(Satir)
End If 

Dim Bag,Conne
Set Bag = New Baglanti
Set Conne = Bag.Connect(Sunucu,VeriTabani,Kullanici,Sifre)

For x=Session("DeathLog") To UBound(Satir)-1

Parca=Split(Satir(x),",")
If Parca(5)>0 And Parca(13)>0 Then
If Trim(Parca(5))="1" Then
Color1="#0099FF"
Else
Color1="#FF0000"
End If
If Trim(Parca(13))="1" Then
Color2="#0099FF"
Else
Color2="#FF0000"
End If

Set ZoneId=Conne.Execute("Select Bz From Zone_info Where Zoneno="&parca(3))

Mesaj = "<b>- <a href=""Karakter-Detay/"&Parca(4)&""" onclick=""pageload('Karakter-Detay/"&Parca(4)&"');return false""><span style=""color:"&Color1&""">" & Parca(4) & "</span></a> >>> <a href=""Karakter-Detay/"&Parca(12)&""" onclick=""pageload('Karakter-Detay/"&Parca(12)&"');return false""><span style=""color:"&Color2&""">" & Parca(12) & "</span></a> Adlý Oyuncuyu Öldürdü. ( "&ZoneId("Bz")&" - "&Parca(0)&":"&Parca(1)&":"&Parca(2)&") -</b><br>"

Response.Write Mesaj

End If

Next

Session("DeathLog") = UBound(Satir)

End If

DeathLog(1,0) = txtfile.size

Ag.Close

End If

DeathLog(0,0) = Now()

Application.Lock
Application("Deathlog") = DeathLog
Application.UnLock

Set txtfile = Nothing

End if

End if

Set objfso = Nothing

%>