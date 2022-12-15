<% 
if Session("durum")="giris_yapmis" then
yuid = session("uye_id")


if request.Form("mesaj_baslik") = "" then
mbaslik = unsubjected
else
mbaslik = guvenlik(request.Form("mesaj_baslik"))
end if

auid = guvenlik(request.querystring("auid"))
mmesaj = guvenmesajyaz(request.form("mesaj_icerik"))
mtarih = Year(date)&"-"&Month(date)&"-"&Day(date)&" "&Hour(now)&":"&Minute(now)&":"&second(now)

SQL="insert into gop_mesajlar (yollayan,alici,mesaj_baslik,mesaj_icerik,mesaj_tarih) values ('"&yuid&"','"&auid&"','"&mbaslik&"','"&mmesaj&"','"&mtarih&"')"
baglanti.execute(SQL)

response.Write "<center>"&sent_message&"</center>"



else
Response.Write "<center>"&notice4&"</center>"
end if
 %>