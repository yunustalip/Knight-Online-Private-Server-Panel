<%
tarihkontrol = Year(date)&"-"&Month(date)&"-"&Day(date)&" "&Hour(now)&":"&Minute(now)-3&":"&second(now)
buguntarih = Year(date)&"-"&Month(date)&"-"&Day(date)
duntarih = Year(date)&"-"&Month(date)&"-"&Day(date)-1

set tarihac = baglanti.execute("SELECT * FROM gop_sayac where sayac_tarih='"& buguntarih &"'")
if tarihac.eof or tarihac.bof then

Baglanti.Execute("insert into gop_sayac (sayac_tarih,sayac_tekil,sayac_cogul) values ('"&buguntarih&"','1','1')")

end if
tarihac.close


set ayarlar = baglanti.execute("select * from gop_sayacayar")

Set rs = server.createobject("ADODB.recordset")
sql="select * from gop_sayac where sayac_tarih='"& buguntarih &"'"
rs.open sql,baglanti,1,3

if request.cookies("gop_sayac")="evet" then
rs("sayac_tekil")=rs("sayac_tekil")+0
else
rs("sayac_tekil")=rs("sayac_tekil")+1
rs.update
response.cookies("gop_sayac")="evet"
end if

session("sayac_tekil") = rs("sayac_tekil")

rs("sayac_cogul")=rs("sayac_cogul")+1
rs.update

session("sayac_cogul") = rs("sayac_cogul")

Set TekilToplami = baglanti.Execute("SELECT SUM(sayac_tekil) AS tekil_toplam FROM gop_sayac")
Set CogulToplami = baglanti.Execute("SELECT SUM(sayac_cogul) AS cogul_toplam FROM gop_sayac")
session("tekil_toplam") = TekilToplami("tekil_toplam")
session("cogul_toplam") = CogulToplami("cogul_toplam")

Set uye_topla = baglanti.Execute("SELECT COUNT(*) as TOPLAM FROM gop_uyeler WHERE uye_son_tarih >= '"&tarihkontrol&"' ORDER BY uye_son_tarih desc;")

session("uye_topla") = uye_topla("TOPLAM")

Set ToplamHaber = baglanti.Execute("SELECT COUNT(vid) AS haber FROM gop_veriler")
Set ToplamOkuma = baglanti.Execute("SELECT SUM(vhit) AS okuma FROM gop_veriler")
session("toplam_haber") = ToplamHaber("haber")
session("toplam_okuma") = ToplamOkuma("okuma")

Set sonuye = baglanti.Execute("SELECT uye_adi, uye_tarih FROM gop_uyeler ORDER BY uye_tarih desc;")
session("son_uye") = sonuye("uye_adi")

%>