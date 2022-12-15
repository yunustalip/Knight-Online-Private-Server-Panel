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
%>
<!--#include file="md5.asp"-->

<%
islem = request.querystring("islem")
if islem = "oku" then oku
if islem = "yeniuye" then yeniuye
if islem = "kategori" then kategori
if islem = "haberler" then haberler
if islem = "hata" then hata
if islem = "bilgilerim" then bilgilerim
if islem = "altkategori" then altkategori
if islem = "uye_guncelle" then uye_guncelle
if islem = "uye_guncelle_islem" then uye_guncelle_islem
if islem = "yorum_gonder" then yorum_gonder
if islem = "uye_islem" then uye_islem
if islem = "uye_kontrol" then uye_kontrol
if islem = "uyeler" then uyeler
if islem = "mesajlarim" then mesajlarim
if islem = "mesaj_oku" then mesaj_oku
if islem = "mesaj_oku_giden" then mesaj_oku_giden
if islem = "mesaj_gonder" then mesaj_gonder
if islem = "mesaj_sil" then mesaj_sil
if islem = "mesaj_giden_sil" then mesaj_giden_sil
if islem = "bilgi_uye" then bilgi_uye
if islem = "tema_degis" then tema_degis
if islem = "etiket" then etiket
if islem = "bilesen" then bilesen
if islem = "" then ana

%>
<% sub ana %>
<!--#include file="../modules/anasayfa.asp"-->
<% end sub %>
<% sub oku %>
<!--#include file="../modules/oku.asp"-->
<% end sub %>
<% sub yeniuye %>
<!--#include file="../includes/uye_kayit.asp"-->
<% end sub %>
<% sub kategori %>
<!--#include file="../includes/kategori.asp"-->
<% end sub %>
<% sub haberler %>
<!--#include file="../modules/haberler.asp"-->
<% end sub %>
<% sub hata %>
<!--#include file="../includes/uye_hata.asp"-->
<% end sub %>
<% sub bilgilerim %>
<!--#include file="../includes/uye_bilgileri.asp"-->
<% end sub %>
<% sub altkategori %>
<!--#include file="../includes/altkategori.asp"-->
<% end sub %>
<% sub uye_guncelle %>
<!--#include file="../includes/uye_guncelle.asp"-->
<% end sub %>
<% sub uye_guncelle_islem %>
<!--#include file="../includes/uye_guncelle_islem.asp"-->
<% end sub %>
<% sub uyeler %>
<!--#include file="../includes/uyeler.asp"-->
<% end sub %>
<% sub mesajlarim %>
<!--#include file="../includes/mesajlarim.asp"-->
<% end sub %>
<% sub mesaj_oku %>
<!--#include file="../includes/mesaj_oku.asp"-->
<% end sub %>
<% sub mesaj_oku_giden %>
<!--#include file="../includes/mesaj_oku_giden.asp"-->
<% end sub %>
<% sub mesaj_gonder %>
<!--#include file="../includes/mesaj_gonder.asp"-->
<% end sub %>
<% sub mesaj_sil %>
<!--#include file="../includes/mesaj_sil.asp"-->
<% end sub %>
<% sub mesaj_giden_sil %>
<!--#include file="../includes/mesaj_sil_giden.asp"-->
<% end sub %>
<% sub bilgi_uye %>
<!--#include file="../includes/uye.asp"-->
<% end sub %>
<% sub yorum_gonder %>
<!--#include file="../includes/yorum_gonder.asp"-->
<% end sub %>
<% sub uye_islem %>
<!--#include file="../includes/uye_islem.asp"-->
<% end sub %>
<% sub uye_kontrol %>
<!--#include file="../includes/uye_kontrol.asp"-->
<% end sub %>
<% sub etiket %>
<!--#include file="../modules/etiketler.asp"-->
<% end sub %>
<% sub tema_degis %>
<!--#include file="../includes/tema_guncelle.asp"-->
<% end sub %>
<% sub bilesen %>
<!--#include file="../includes/bilesenler.asp"-->
<% end sub %>
<% 'Bölümleme bitimi %>
<%
if google = "1" then
Response.Write "<br /><p align=""center"">"
Response.Write ("" & _
vbCrLf & "<script type=""text/javascript""><!--" & _
vbCrLf & "google_ad_client = """&google_code&""";" & _
vbCrLf & "google_alternate_ad_url = ""http://www.joomlasp.com/images/reklam.png"";" & _
vbCrLf & "google_ad_width = 468;" & _
vbCrLf & "google_ad_height = 60;" & _
vbCrLf & "google_ad_format = ""468x60_as"";" & _
vbCrLf & "google_ad_type = ""text"";" & _
vbCrLf & "google_ad_channel =""7188969507"";" & _
vbCrLf & "google_color_border = ""6699CC"";" & _
vbCrLf & "google_color_bg = ""003366"";" & _
vbCrLf & "google_color_link = ""FFFFFF"";" & _
vbCrLf & "google_color_url = ""AECCEB"";" & _
vbCrLf & "google_color_text = ""AECCEB"";" & _
vbCrLf & "//--></script>" & _
vbCrLf & "<script type=""text/javascript"" src=""http://pagead2.googlesyndication.com/pagead/show_ads.js""></script>")
Response.Write "</p>"
end if
%>