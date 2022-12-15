<%response.charset="iso-8859-9"

gun=WeekdayName(weekday(date))
saat=hour(now)
dakika=minute(now)

select case gun
case "pazartesi" 
if saat<10 Then
if saat=>6 and dakika>=30 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
End If
elseif saat>=10 and saat<15 Then
dj="Yalçın"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=7"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=15 and saat<19 Then
dj="Özlem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=6"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=19 and saat<=24  Then
dj="Erkan"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=9"
djresim="http://www.superfm.com.tr/images/Icerik/125_20100606_075734.JPG"
End If
case "Salı"
if saat<10 Then
if saat=>6 and dakika>=30 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
End If
elseif saat>=10 and saat<14 Then
dj="Yalçın"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=7"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=14 and saat<18 Then
dj="Özlem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=6"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=18 and saat<20 Then
dj="Tolga"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=8"
djresim="http://www.superfm.com.tr/images/Icerik/tolga_orta_20091106_080907.JPG"
elseif saat>=20 and saat<=24  Then
dj="Erkan"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=9"
djresim="http://www.superfm.com.tr/images/Icerik/125_20100606_075734.JPG"
End If
case "Çarşamba"
if saat<10 Then
if saat=>6 and dakika>=30 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
End If
elseif saat>=10 and saat<14 Then
dj="Yalçın"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=7"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=14 and saat<18 Then
dj="Özlem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=6"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=18 and saat<20 Then
dj="Tolga"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=8"
djresim="http://www.superfm.com.tr/images/Icerik/tolga_orta_20091106_080907.JPG"
elseif saat>=20 and saat<=24 Then
dj="Erkan"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=9"
djresim="http://www.superfm.com.tr/images/Icerik/125_20100606_075734.JPG"
End If
case "Perşembe"
if saat<10 Then
if saat=>6 and dakika>=30 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
End If
elseif saat>=10 and saat<14 Then
dj="Yalçın"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=7"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=14 and saat<18 Then
dj="Özlem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=6"
djresim="http://www.superfm.com.tr/images/Icerik/ozlem_orta_20091106_080940.JPG"
elseif saat>=18 and saat<20 Then
dj="Tolga"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=8"
djresim="http://www.superfm.com.tr/images/Icerik/tolga_orta_20091106_080907.JPG"
elseif saat>=20 and saat<=24 Then
dj="Erkan"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=9"
djresim="http://www.superfm.com.tr/images/Icerik/125_20100606_075734.JPG"
End If
case "Cuma"
if saat<10 Then
if saat=>6 and dakika>=30 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
End If
elseif saat>=10 and saat<14 Then
dj="Yalçın"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=7"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=14 and saat<18 Then
dj="Özlem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=6"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=18 and saat<20 Then
dj="Tolga"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=8"
djresim="http://www.superfm.com.tr/images/Icerik/tolga_orta_20091106_080907.JPG"
elseif saat>=20 and saat<22 Then
dj="Erkan"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=9"
djresim="http://www.superfm.com.tr/images/Icerik/125_20100606_075734.JPG"
elseif saat>=22 and saat<=24 Then
dj="SUPER HAFTASONU"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=14"
djresim="http://www.superfm.com.tr/images/Icerik/super_haftasonu_tolga_20100430_091331.JPG"
End If
case "Cumartesi"
if saat>=9 and saat<13 Then
dj="Duygu"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=10"
djresim="http://www.superfm.com.tr/images/Icerik/duygu_orta_20091106_080819.JPG"
elseif saat>=13 and saat<18 Then
dj="Özlem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=6"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=18 and saat<20 Then
dj="SUPER 20"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=13"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=20 and saat<22 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
elseif saat>=22 and saat<=24  Then
dj="SUPER HAFTASONU"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=14"
djresim="http://www.superfm.com.tr/images/Icerik/super_haftasonu_tolga_20100430_091331.JPG"
End If
case "Pazar"
if saat>=9 and saat<13 Then
dj="Yalçın"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=7"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=13 and saat<15 Then
dj="SUPER 20"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=13"
djresim="http://www.superfm.com.tr/images/Icerik/yalcin_orta_20091106_080924.JPG"
elseif saat>=15 and saat<19 Then
dj="Gizem"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=12"
djresim="http://www.superfm.com.tr/images/Icerik/gizem_orta_20091106_080758.JPG"
elseif saat>=19 and saat<=23 Then
dj="Erkan"
djlink="http://www.superfm.com.tr/programlar.asp?b=detay&ID=9"
djresim="http://www.superfm.com.tr/images/Icerik/125_20100606_075734.JPG"
End If
case else
end select


Response.Write "&dj="&dj&"&djresim="&djresim


%>
