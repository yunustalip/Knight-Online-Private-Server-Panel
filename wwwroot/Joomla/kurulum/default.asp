<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<meta name="author" content="JoomlASP | Hasan Emre Asker">
<meta name="keywords" content="JoomlASP, Joomla, MySQL, ASP, Active Server Page, ASP Portal, JoomlASP temalari, JoomlASP modülleri, JoomlASP bilesenleri, Site içerik yönetimi, JoomlASP Portali">
<meta name="description" content="JoomlASP - Gelisime Açik Site Içerik Yönetimi">
<link href="../favicon.ico" rel="JoomlASP" />
</head>

<style type="text/css">
<!--
body,td,th {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
}
.style2 {
	color: #006600;
	font-weight: bold;
}
-->
</style>

<!--#include file="../functions/db.asp"-->
<%
Set Baglanti= Server.CreateObject("ADODB.Connection")
Baglanti.open "driver={SQL Server}; SERVER="&mysql_server&"; UID="&mysql_user&"; pwd="&mysql_pass&";db="&mysql_db&"; option = 999999"

%>

  <table width="750" border="0" align="center">
    <tr>
      <td style="border:solid 1px; border-color:#000000;"><img src="images/joomlasp_kur.png" width="750" height="108"></td>
    </tr><tr>
      <td style="border:solid 1px; border-color:#000000;"><div align="center">
<%
islem = request.querystring("islem")
if islem = "sqlkontrol" then sqlkontrol
if islem = "adim1" then adim1
if islem = "adim2" then adim2
if islem = "adim3" then adim3
if islem = "adim4" then adim4
if islem = "adim5" then adim5

if islem = "" then default

sub default
%>
    
        <p>&nbsp;</p>
        <p>--- Lütfen MYSQL Bilgilerinizi functions/db.asp içine giriniz.---</p>
        <p class="style2">&lt;&lt; Bağlantınız Kuruldu &gt;&gt;</p>
        <p><a href="?islem=adim1">Kuruluma Devam Et &gt;&gt;</a></p>
      </div>
<%
end sub

sub adim1

Response.Write "<br><br>Tablolar Oluşturuluyor.<br><br><table><tr><td align=""left"">"



sql = "CREATE TABLE gop_adminmenua ( "
sql = sql & "amid INT default null ,"
sql = sql & "amadi varchar(55) default NULL, "
sql = sql & "amar varchar(255) default NULL, "
sql = sql & "amka varchar(50) default NULL, "
sql = sql & "amord INT default NULL, "
sql = sql & ") "
baglanti.Execute sql
Response.Write "* Admin Menu Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_adminmenub ("
sql = sql & "ambid int(11) NOT NULL ,"
sql = sql & "ambadi varchar(55) default NULL,"
sql = sql & "amblink varchar(255) default NULL,"
sql = sql & "amka varchar(50) default NULL,"
sql = sql & "PRIMARY KEY  (ambid)"

baglanti.Execute sql
Response.Write "* Admin Alt Menu Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_anakat ("
sql = sql & "ankid int(11) NOT NULL ,"
sql = sql & "ankadi varchar(255) default NULL,"
sql = sql & "ankgoster int(11) NOT NULL default 0,"
sql = sql & "ankorder int(11) NOT NULL default 1,"
sql = sql & "ankbilgi varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (ankid)"

baglanti.Execute sql
Response.Write "* Kategori Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_anketcevap ("
sql = sql & "no int(10) NOT NULL ,"
sql = sql & "poll_id int(10) default NULL,"
sql = sql & "answer_id int(10) default NULL,"
sql = sql & "answer varchar(200) default NULL,"
sql = sql & "no_votes int(10) NOT NULL default 0,"
sql = sql & "PRIMARY KEY  (no)"

baglanti.Execute sql
Response.Write "* Anket Cevapları Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_anketip ("
sql = sql & "no int(10) NOT NULL ,"
sql = sql & "poll_id_ip int(10) NOT NULL default 0,"
sql = sql & "ip varchar(25) default 255.255.255.255,"
sql = sql & "PRIMARY KEY  (no)"

baglanti.Execute sql
Response.Write "* Anket IP Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_anketsoru ("
sql = sql & "no int(10) NOT NULL ,"
sql = sql & "id int(10) default NULL,"
sql = sql & "title varchar(200) default NULL,"
sql = sql & "active int(1) NOT NULL default 0,"
sql = sql & "votes int(10) NOT NULL default 0,"
sql = sql & "expiration_start date default NULL COMMENT start date,"
sql = sql & "expiration_end date default NULL COMMENT end date,"
sql = sql & "PRIMARY KEY  (no)"

baglanti.Execute sql
Response.Write "* Anket Soru Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_ayarlar ("
sql = sql & "ayar_id int(11) NOT NULL ,"
sql = sql & "vsayi int(11) default NULL,"
sql = sql & "vsutun int(11) default NULL,"
sql = sql & "vkarakter int(11) default NULL,"
sql = sql & "vyorum varchar(11) default ,"
sql = sql & "siteadi varchar(255) default NULL,"
sql = sql & "admin_mail varchar(255) default NULL,"
sql = sql & "meta_desc mediumtext,"
sql = sql & "meta_key mediumtext,"
sql = sql & "aspjpeg varchar(11) NOT NULL default hayir,"
sql = sql & "siteadres varchar(255) default NULL,"
sql = sql & "kresim int(11) default NULL,"
sql = sql & "google int(11) default NULL,"
sql = sql & "google_code varchar(255) default NULL,"
sql = sql & "yenileme int(3) NOT NULL default 240,"
sql = sql & "eklenti_k varchar(25) NOT NULL default main_page,"
sql = sql & "dilayari int(11) NOT NULL default 1,"
sql = sql & "PRIMARY KEY  (ayar_id)"

baglanti.Execute sql
Response.Write "* Ayarlar Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_download ("
sql = sql & "down_id int(11) NOT NULL ,"
sql = sql & "down_adi varchar(99) default NULL,"
sql = sql & "down_bilgi mediumtext,"
sql = sql & "down_link varchar(255) default NULL,"
sql = sql & "down_hit int(11) default NULL,"
sql = sql & "dkid int(11) default NULL,"
sql = sql & "down_resim varchar(50) default NULL,"
sql = sql & "down_onay int(11) NOT NULL default 0,"
sql = sql & "PRIMARY KEY  (down_id)"

baglanti.Execute sql
Response.Write "* Download Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_download_kat ("
sql = sql & "dkid int(11) NOT NULL ,"
sql = sql & "dk_adi varchar(99) default NULL,"
sql = sql & "PRIMARY KEY  (dkid)"

baglanti.Execute sql
Response.Write "* Download Kategori Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_eklentiler ("
sql = sql & "id int(11) NOT NULL ,"
sql = sql & "eklenti_adi varchar(55) default NULL,"
sql = sql & "eklenti mediumtext,"
sql = sql & "eklenti_yazar varchar(55) default NULL,"
sql = sql & "eklenti_mail varchar(255) default NULL,"
sql = sql & "eklenti_k varchar(55) default NULL,"
sql = sql & "eklenti_web mediumtext,"
sql = sql & "eklenti_kaldir mediumtext,"
sql = sql & "PRIMARY KEY  (id)"

baglanti.Execute sql
Response.Write "* Bileşenler Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_galeri ("
sql = sql & "resid int(11) NOT NULL ,"
sql = sql & "galid int(11) default NULL,"
sql = sql & "resadi varchar(50) default NULL,"
sql = sql & "resresim varchar(255) default NULL,"
sql = sql & "rhit int(11) NOT NULL default 0,"
sql = sql & "PRIMARY KEY  (resid)"

baglanti.Execute sql
Response.Write "* Galeri Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_galerikat ("
sql = sql & "galid int(11) NOT NULL ,"
sql = sql & "galkat varchar(25) default NULL,"
sql = sql & "PRIMARY KEY  (galid)"

baglanti.Execute sql
Response.Write "* Galeri Kategori Tablosu Oluşturuldu<br>"


sql = "CREATE TABLE gop_group ("
sql = sql & "gid int(11) NOT NULL ,"
sql = sql & "gadi varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (gid)"

baglanti.Execute sql
Response.Write "* Gruplar Tablosu Oluşturuldu<br>"

sql = "CREATE TABLE gop_iletisim ("
sql = sql & "id int(11) NOT NULL ,"
sql = sql & "adi varchar(50) default NULL,"
sql = sql & "mail varchar(255) default NULL,"
sql = sql & "yas int(2) default NULL,"
sql = sql & "tel varchar(25) default NULL,"
sql = sql & "mesaj mediumtext,"
sql = sql & "tarih datetime NOT NULL default 2008-01-01 12:00:00,"
sql = sql & "PRIMARY KEY  (id)"

baglanti.Execute sql
Response.Write "* İletişim Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_kat ("
sql = sql & "katid int(11) NOT NULL ,"
sql = sql & "katadi varchar(255) default NULL,"
sql = sql & "katbilgi varchar(255) default NULL,"
sql = sql & "ankid int(11) default NULL,"
sql = sql & "PRIMARY KEY  (katid)"

baglanti.Execute sql
Response.Write "* Alt Kategori Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_language ("
sql = sql & "lang_id int(11) NOT NULL ,"
sql = sql & "language mediumtext,"
sql = sql & "lang_adi varchar(25) default NULL,"
sql = sql & "lang_yazar varchar(50) default NULL,"
sql = sql & "lang_mail varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (lang_id)"

baglanti.Execute sql
Response.Write "* Dil Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_linkler ("
sql = sql & "link_id int(11) NOT NULL ,"
sql = sql & "link_adi varchar(255) default NULL,"
sql = sql & "link_aciklama mediumtext,"
sql = sql & "link_onay int(11) default 0,"
sql = sql & "link_gosterim int(11) default 0,"
sql = sql & "link_url varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (link_id)"

baglanti.Execute sql
Response.Write "* Linkler Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_menu ("
sql = sql & "m_id int(11) NOT NULL ,"
sql = sql & "m_adi varchar(25) default NULL,"
sql = sql & "m_link varchar(255) default NULL,"
sql = sql & "m_order int(11) NOT NULL default 99,"
sql = sql & "m_ust int(11) default NULL,"
sql = sql & "m_yan int(11) default NULL,"
sql = sql & "PRIMARY KEY  (m_id)"

baglanti.Execute sql
Response.Write "* Menuler Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_mesajlar ("
sql = sql & "mesaj_id int(11) NOT NULL ,"
sql = sql & "yollayan int(11) default NULL,"
sql = sql & "alici int(11) default NULL,"
sql = sql & "mesaj_baslik varchar(50) default NULL,"
sql = sql & "mesaj_icerik mediumtext,"
sql = sql & "mesaj_tarih datetime default NULL,"
sql = sql & "mesaj_sil int(1) NOT NULL default 0,"
sql = sql & "mesaj_okundu int(1) NOT NULL default 0,"
sql = sql & "mesaj_gsil int(1) NOT NULL default 0,"
sql = sql & "PRIMARY KEY  (mesaj_id)"

baglanti.Execute sql
Response.Write "* Mesajlar Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_modules ("
sql = sql & "modul_id int(11) NOT NULL ,"
sql = sql & "modul_adi varchar(99) default NULL,"
sql = sql & "modul_icerik mediumtext,"
sql = sql & "modul_yer varchar(11) NOT NULL default sol,"
sql = sql & "modul_sira int(11) NOT NULL default 1,"
sql = sql & "modul_izin int(11) NOT NULL default 0,"
sql = sql & "modul_yazar varchar(50) default NULL,"
sql = sql & "modul_mail varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (modul_id)"

baglanti.Execute sql
Response.Write "* Modüller Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_reklam ("
sql = sql & "rid int(11) NOT NULL ,"
sql = sql & "rgoster int(11) NOT NULL default 0,"
sql = sql & "rresim mediumtext,"
sql = sql & "hit int(11) NOT NULL default 0,"
sql = sql & "rlink varchar(255) default NULL,"
sql = sql & "radi varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (rid)"

baglanti.Execute sql
Response.Write "* Reklamlar Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_sayac ("
sql = sql & "say_id int(10) NOT NULL ,"
sql = sql & "sayac_tekil int(10) NOT NULL default 0,"
sql = sql & "sayac_cogul int(10) NOT NULL default 0,"
sql = sql & "sayac_tarih date NOT NULL default 0000-00-00,"
sql = sql & "PRIMARY KEY  (say_id)"

baglanti.Execute sql
Response.Write "* İstatistik Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_sayacayar ("
sql = sql & "id int(11) NOT NULL ,"
sql = sql & "btekil int(11) NOT NULL default 0,"
sql = sql & "bcogul int(11) NOT NULL default 0,"
sql = sql & "toplamc int(11) NOT NULL default 0,"
sql = sql & "toplamt int(11) NOT NULL default 0,"
sql = sql & "dtekil int(11) NOT NULL default 0,"
sql = sql & "dcogul int(11) NOT NULL default 0,"
sql = sql & "aktifuye int(11) NOT NULL default 0,"
sql = sql & "okunma int(11) NOT NULL default 0,"
sql = sql & "ip int(11) NOT NULL default 0,"
sql = sql & "online int(11) NOT NULL default 0,"
sql = sql & "veri int(11) NOT NULL default 0,"
sql = sql & "sonuye int(11) NOT NULL default 0,"
sql = sql & "PRIMARY KEY  (id)"

baglanti.Execute sql
Response.Write "* İstatistik Ayarlar Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_sayfa ("
sql = sql & "sayfaid int(11) NOT NULL ,"
sql = sql & "sayfa_baslik varchar(255) default NULL,"
sql = sql & "sayfa_icerik mediumtext,"
sql = sql & "sayfa_hit int(11) NOT NULL default 0,"
sql = sql & "sayfa_tarih datetime NOT NULL default 0000-00-00 00:00:00,"
sql = sql & "PRIMARY KEY  (sayfaid)"

baglanti.Execute sql
Response.Write "* Sayfalar Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_uyeler ("
sql = sql & "uye_id int(11) NOT NULL ,"
sql = sql & "gid int(11) NOT NULL default 3,"
sql = sql & "uye_adi varchar(25) default NULL,"
sql = sql & "uye_sifre varchar(50) default NULL,"
sql = sql & "uye_mail varchar(35) default NULL,"
sql = sql & "uye_isim varchar(25) default NULL,"
sql = sql & "uye_soyisim varchar(25) default NULL,"
sql = sql & "uye_website varchar(50) default NULL,"
sql = sql & "uye_ulke varchar(25) default NULL,"
sql = sql & "uye_sehir varchar(25) default NULL,"
sql = sql & "uye_msn varchar(50) default NULL,"
sql = sql & "uye_icq varchar(15) default NULL,"
sql = sql & "uye_aol varchar(50) default NULL,"
sql = sql & "uye_yahoo varchar(50) default NULL,"
sql = sql & "uye_skype varchar(50) default NULL,"
sql = sql & "uye_avatar varchar(100) default NULL,"
sql = sql & "uye_tarih datetime NOT NULL default 0000-00-00 00:00:00,"
sql = sql & "uye_son_tarih datetime NOT NULL default 0000-00-00 00:00:00,"
sql = sql & "uye_dil int(11) NOT NULL default 1,"
sql = sql & "uye_ip varchar(25) NOT NULL default 255.255.255.255,"
sql = sql & "PRIMARY KEY  (uye_id)"

baglanti.Execute sql
Response.Write "* Üyeler Tablosu Oluşturuldu<br>"



sql = "CREATE TABLE gop_veriler ("
sql = sql & "vid int(11) NOT NULL ,"
sql = sql & "vbaslik varchar(255) default NULL,"
sql = sql & "vgoster int(11) default NULL,"
sql = sql & "katid int(11) default NULL,"
sql = sql & "vicerik mediumtext,"
sql = sql & "vhit int(11) NOT NULL default 0,"
sql = sql & "vresim varchar(255) default NULL,"
sql = sql & "uye_id int(11) default NULL,"
sql = sql & "vtarih date NOT NULL default 0000-00-00,"
sql = sql & "vetiket mediumtext NOT NULL,"
sql = sql & "PRIMARY KEY  (vid)"

baglanti.Execute sql
Response.Write "* Veriler Tablosu Oluşturuldu<br>"




sql = "CREATE TABLE gop_yorumlar ("
sql = sql & "yorum_id int(11) NOT NULL ,"
sql = sql & "vid int(11) default NULL,"
sql = sql & "yorum mediumtext,"
sql = sql & "yorum_tarih datetime default NULL,"
sql = sql & "uye_id int(11) default NULL,"
sql = sql & "yorum_onay int(11) default 0,"
sql = sql & "PRIMARY KEY  (yorum_id)"

baglanti.Execute sql
Response.Write "* Yorumlar Tablosu Oluşturuldu<br>"
Response.Write "<br><br><center><a href=""?islem=adim2"">Devam Et >> </a></center></td></tr></table>"
end sub

sub adim2

Response.Write "<br><br>Veriler giriliyor.<br><br><table><tr><td align=""left"">"
Response.Write "* Admin ana menüsü işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (1,Response.Write general_settings,genel_ayarlar.png,genel_ayarlar,1);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (2,Response.Write menu_settings,menu_islemleri.png,menu_islemleri,3);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (3,Response.Write category_settings,kategori_islemleri.png,kategori_islemleri,4);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (4,Response.Write data_settings,veri_islemleri.png,veri_islemleri,5);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (5,Response.Write addition_settings,modul_islemleri.png,eklenti_islemleri,6);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (6,Response.Write members_settings,uye_islemleri.png,uye_islemleri,7);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (7,Response.Write link_settings,link_islemleri.png,link_islemleri,8);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (8,Response.Write download_settings,download_islemleri.png,download_islemleri,9);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (9,Response.Write gallery_settings,galeri_islemleri.png,galeri_islemleri,10);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (10,Response.Write poll_settings,anket_yonetimi.png,anket_islemleri,11);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (11,Response.Write advertising_settings,reklam_yonetimi.png,reklam_islemleri,12);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (12,Response.Write communication_settings,iletisim_islemleri.png,iletisim_islemleri,13);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (13,Response.Write language_settings,dil_islemleri.png,dil_islemleri,2);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (14,Response.Write page_settings,veri_islemleri.png,sayfa_islemleri,5);"
Baglanti.execute "INSERT INTO gop_adminmenua VALUES (15,Response.Write statistics ,sayac_ayarlari.asp,sayac_islem,14);"
Response.Write "<b>* İşlendi.</b><br>"


Response.Write "* Admin menüsü işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (1,Response.Write site_settings,site_ayarlari.asp,genel_ayarlar);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (2,Response.Write menu,menuler.asp,menu_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (3,Response.Write add,menuler.asp?islem=menu_ekle,menu_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (4,Response.Write category,kategoriler.asp,kategori_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (5,Response.Write add,kategoriler.asp?act=ekle,kategori_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (6,Response.Write sub_category,altkategoriler.asp,kategori_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (7,Response.Write add,altkategoriler.asp?islem=yeni,kategori_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (8,Response.Write data,veriler.asp,veri_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (9,Response.Write add,veriler.asp?islem=ekle,veri_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (10,Response.Write comment,yorumlar.asp,veri_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (11,Response.Write review_approve,yorumlar.asp?islem=bekleyen,veri_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (12,Response.Write modules,moduller.asp,eklenti_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (13,Response.Write add_manuel_module,moduller.asp?islem=modul_manuel_ekle,eklenti_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (14,Response.Write members,uyeler.asp,uye_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (15,Response.Write add,uyeler.asp?islem=ekle,uye_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (16,Response.Write links,linkler.asp,link_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (17,Response.Write add,linkler.asp?act=ekle,link_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (18,Response.Write approve,linkler.asp?act=onay,link_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (19,Response.Write category,downloadlar.asp?islem=kategoriler,download_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (20,Response.Write downloads,downloadlar.asp,download_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (21,Response.Write add,downloadlar.asp?islem=ekle,download_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (22,Response.Write approve,downloadlar.asp?islem=onay,download_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (23,Response.Write category,galeri.asp,galeri_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (24,Response.Write pictures,galeri.asp?islem=resimler,galeri_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (25,Response.Write polls,anket.asp,anket_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (26,Response.Write add,anket.asp?sub=addnew,anket_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (27,Response.Write advertisement,reklamlar.asp,reklam_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (28,Response.Write inbox,iletisim.asp,iletisim_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (29,Response.Write add_component,bilesenler.asp,eklenti_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (30,Response.Write language,diller.asp,dil_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (31,Response.Write database_backup,db_yedekle.asp,genel_ayarlar);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (32,Response.Write pages,sayfalar.asp,sayfa_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (33,Response.Write add_page,sayfalar.asp?islem=ekle,sayfa_islemleri);"
Baglanti.execute "INSERT INTO gop_adminmenub VALUES (34,Response.Write statistics ,sayac_ayarlari.asp,sayac_islem);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Katgori işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_anakat VALUES (1,Haberler,1,1,Sitemiz hakkindaki haberlerin bulundugu kategori);"
Baglanti.execute "INSERT INTO gop_anakat VALUES (2,Eklentiler,1,2,Eklentilerim);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Anketler işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (1,3,8,Asp,552);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (2,3,9,Php,110);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (3,3,10,Java Servlet,20);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (4,3,11,Cold Fusion,29);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (5,4,12,Kesinlikle Yapılmalı,11);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (6,4,13,Pek Gerek Yok,1);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (7,4,14,Hiç Gerek Yok,0);"
Baglanti.execute "INSERT INTO gop_anketcevap VALUES (8,4,15,Karasızım,1);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Anket Ip işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_anketip VALUES (1,4,\78.186.8.173\);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Anket Soruları işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_anketsoru VALUES (1,3,Yazılım Diliniz ?,0,710,2008-04-05,2009-05-05);"
Baglanti.execute "INSERT INTO gop_anketsoru VALUES (2,4,JoomlASP için Forum Sistemi Yapılsın mı?,1,13,2009-02-03,2009-05-05);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Site ayarları işleniyor<br>"
Baglanti.execute "INSERT INTO gop_ayarlar VALUES (1,10,1,600,goster,JoomlASP 1.0.3 | Gelisime Açik Site Yönetim Sistemi,admin@joomlasp.com,JoomlASP - Gelisime Açik Site Içerik Yönetimi,JoomlASP, Joomla, MySQL, ASP, Active Server Page, ASP Portal, JoomlASP temalari, JoomlASP modülleri, JoomlASP bilesenleri, Site içerik yönetimi, JoomlASP Portali,evet,http://www.joomlasp.com/demo,1,1,pub-3870177390472952,240,main_page,1);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Downloadlar işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_download VALUES (1,Test1,Test download bilgisi,http://www.joomlasp.com,500,10,,1);"
Baglanti.execute "INSERT INTO gop_download VALUES (2,Test2,Download test2 bilgisi,http://www.joomlasp.com,712,10,,1);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Download Kategorileri işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_download_kat VALUES (1,Test);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Galeri işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_galeri VALUES (1,1,Kis,http://www.joomlasp.com/demo/galeri/kis.jpg,9);"
Baglanti.execute "INSERT INTO gop_galeri VALUES (2,1,Kis - 2,http://www.joomlasp.com/demo/galeri/kis2.jpg,6);"
Baglanti.execute "INSERT INTO gop_galeri VALUES (3,1,Manzara,http://www.joomlasp.com/demo/galeri/galeri1(1).jpg,24);"
Baglanti.execute "INSERT INTO gop_galeri VALUES (4,1,Kars,http://www.joomlasp.com/demo/galeri/kars_park_1_b(1).jpg,13);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Galeri Kategorileri işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_galerikat VALUES (1,Genel);"
Baglanti.execute "INSERT INTO gop_galerikat VALUES (2,Ask & Sevgi);"
Baglanti.execute "INSERT INTO gop_galerikat VALUES (3,Korku);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Gruplar işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_group VALUES (1,Administrator);"
Baglanti.execute "INSERT INTO gop_group VALUES (2,Uye);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Alt Kategoriler işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_kat VALUES (1,Haberler,JoomlASP Haberleri,4);"
Baglanti.execute "INSERT INTO gop_kat VALUES (2,Bilesenler,,5);"
Baglanti.execute "INSERT INTO gop_kat VALUES (3,Modüller,,5);"
Response.Write "<b>* İşlendi.</b><br>"


Response.Write "* Linkler işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_linkler VALUES (1,JoomlASP,JoomlASP Resmi sitesidir. Buradan JoomlASP sisteminizi gelistirmek için eklentileri indirip kurabilirsiniz. Ayrica Resmi Destek Sitesidir (RDS), bunun haricinde bulunan bir resmi sitesi kesinlikle mevcut degildir. Baska yerlerden edindiginiz eklentilerden JoomlASP sorumlu degildir.,1,363,www.joomlasp.com);"
Baglanti.execute "INSERT INTO gop_linkler VALUES (2,Gopca.Net,JoomlASP ile hazirlanmis bir Teknohaber sitesi. Sürekli güncel ve joomlASP\nin tüm özelliklerini sonuna kadar kullanmaktadir. JoomlASP ile neler yapabileceginizi görebilmeniz için mutlaka giriniz.,1,303,www.gopca.net);"


Response.Write "* Menüler işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_menu VALUES (1,Ana Sayfa,default.asp,1,1,1);"
Baglanti.execute "INSERT INTO gop_menu VALUES (2,Iletisim,default.asp?islem=bilesen&component=iletisim,6,0,1);"
Baglanti.execute "INSERT INTO gop_menu VALUES (3,Download,default.asp?islem=bilesen&component=download,3,0,1);"
Baglanti.execute "INSERT INTO gop_menu VALUES (4,Linkler,default.asp?islem=bilesen&component=linkler,4,0,1);"
Baglanti.execute "INSERT INTO gop_menu VALUES (5,Resim Galerisi,default.asp?islem=bilesen&component=galeri,9,0,1);"
Baglanti.execute "INSERT INTO gop_menu VALUES (6,Hakkımızda,default.asp?islem=bilesen&component=sayfa_sistemi&sayfaid=1&sayfa_adi=JoomlASP,4,0,1);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Mesajlar işleniyor<br>."
Baglanti.execute "INSERT INTO gop_mesajlar VALUES (1,1,1,Hoşgeldiniz...,JoomlASP kurulumunu başarıyla tamamlamış bulunmaktasınız. Kullanım ve Destek için JoomlASP Resmi sitesi olan www.joomlasp.com adresinizi ziyaret edebilirsiniz.,2009-02-06 16:52:29,0,0,0);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Reklamlar işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_reklam VALUES (1,1,http://www.joomlasp.com/images/reklam.png,39334,http://www.joomlasp.com,JoomlASP);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* İstatistik bilgileri işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_sayac VALUES (1,0,0,2009-02-06);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* İstatistik Ayarları işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_sayacayar VALUES (1,1,1,1,1,1,1,1,1,1,1,1,1);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Sayfalar işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_sayfa VALUES (1,JoomlASP,JoomlASP, Asp (Active Server Page) dilinde yazilmis olan ve birebir Joomla özeliklerini kapsayan ve gelisime açik bir cms sistemidir. Kisisel web sitelerinden, özel firma sitelerine kadar bir siteyi rahatlikla yapabileceginiz, kolay, modüler, mysql db sistemi hizli bir sekilde islem yapabilen bir cms sistemidir.<br><br>\r\n \r\nSistemin yazilis amaci Php dilinin bir çok kisiye basit gelmesine karsin Asp dili ile internet programciligina baslayan kisilerde Php dilinden uzak kalma ve Php kodlarini çözememe gibi durumlar için hazirlanmistir. Asp dilini uzan yillardir kullanan JoomlASP yapimcisi Hasan Emre ASKER, Php dilinden anlamayan ve karmasik yapiya sahip oldugunu düsünen kisiler için kullanimi Joomlaya göre daha kolay olan bu CMS sistemini meydana getirmistir.<br><br>\r\n\r\nSistemimiz gelistirilmeye açik bir yapiya sahip oldugundan ve Joomla temalari ile birebir uyumluluk sergilediginden begendiginiz Joomla temasini basit bir kaç editleme ile JoomlASP\ye uyarlayabilirsiniz. Eger sizler bunu yapacak vaktimiz yok ben su temanin JoomlASP versiyonunu istiyorum derseniz, site editörlerimiz tarafindan istediginiz bir Joomla temasi JoomlASP sistemine çevrilip temalar bölümünde yayinlanacaktir.<br><br> \r\n\r\nYine sistem için modül yazimini kolaylastirmak için site yapimizi basitlestirmis ve asp dilinden çok az anlayan birinin bile kendine göre küçük modüller yazabilecegine inancimiz tamdir.<br><br> \r\n\r\nSitemiz haricinde yayinlanan JoomlASP bilesen ve Modüllerinden lütfen uzak durunuz. Yazilmis olan Modül, Bilesen ve Temalar sitemiz editörleri tarafindan incelenip açiklarinin var olup olmadigi kontrol edildikten sonra yayina hazir halde sizlere sunulacaktir. Bu yüzden lütfen disardan temin ettiginiz Bilesen, Modül ve Temalari kullanmayiniz. ,546,2008-10-25 10:00:00);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Üyeler işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_uyeler VALUES (1,1,Admin,e10adc3949ba59abbe56e057f20f883e,emre06@hotmail.com.tr,Hasan Emre,Asker,www.joomlasp.com,Türkiye,Kars,emre06@hotmail.com.tr,73974312,,,,,2007-07-20 04:15:46,2009-02-06 15:26:43,1,);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Veriler işleniyor.<br>"
Baglanti.execute "INSERT INTO gop_veriler VALUES (1,JoomlASPye Hosgeldiniz.,1,1,<p>\r\n&nbsp;&nbsp;Son hızla devam eden JoomlASP v1.0.2 Beta Demo Sitemiz&nbsp;tam&nbsp;kapasiteli&nbsp;olarak hizmete girmistir. Sitemizi zitaret edip yeni versiyon hakkinda bilgiler edinebilirsiniz. Unutulmamalidir ki beta versiyonlarda hatalar bulunabilir. Bu hatalari l&uuml;tfen emre06@hotmail.com.tr adli&nbsp;msn adresinden veya info@joomlasp.com adresine mail atarak bilgilendirebiliriniz. \r\n</p>\r\n<p>\r\n&nbsp;&nbsp;Sitemizin tamamen &uuml;cretsiz kullanilabilmesi ve gelisimine katkida bulunmak i&ccedil;in l&uuml;tfen Google Reklam alanlarini kaldirmayiniz. Gerekli bilgiler beni_oku.txt dosyasinin i&ccedil;indedir. \r\n</p>\r\n<p>\r\n&nbsp;\r\n</p>\r\n,0,,1,2007-11-05,JoomlASP, Google, Reklam);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "<br><br><center><a href=""?islem=adim3"">Devam Et >> </a></center></td></tr></table>"
end sub

sub adim3
kurulumadi = "dil.xml"
set xmlDoc = createObject("Microsoft.XMLDOM")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
dosya = Server.MapPath(kurulumadi)
xmlDoc.load (dosya)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatası: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "diladi" then
		diladi = entry.text
        elseif entry.tagname = "aspcode" then 
		aspcode = entry.text
		elseif entry.tagname = "dilyazar" then 
		dilyazar = entry.text
		elseif entry.tagname = "mail" then 
		mail = entry.text
		end if
    next
end If

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_language"
rs.open SQL,baglanti,1,3
rs.addnew
rs("lang_adi") = diladi
rs("language") = aspcode
rs("lang_yazar") = dilyazar
rs("lang_mail") = mail
rs.update

Response.Write "<br><br><center>Dil Dosyası Yüklendi</center><br><br>"

Response.Write "<br><br><center><a href=""?islem=adim4"">Devam Et >> </a></center>"
end sub

sub adim4

for zzz="1" to "7"

kurulumadi = "bilesen"&zzz&".xml"
set xmlDoc = createObject("Microsoft.XMLDOM")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
dosya = Server.MapPath(kurulumadi)
xmlDoc.load (dosya)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatası: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "eklenti_adi" then
		eklenti_adi = entry.text
        elseif entry.tagname = "eklenti_k" then 
		eklenti_k = entry.text
		elseif entry.tagname = "eklenti_yazar" then 
		eklenti_yazar = entry.text
		elseif entry.tagname = "eklenti_mail" then 
		eklenti_mail = entry.text
		elseif entry.tagname = "eklenti_web" then 
		eklenti_web = entry.text
		elseif entry.tagname = "sqlcode" then 
		sqlcode = entry.text
		elseif entry.tagname = "aspcode" then 
		aspcode = entry.text
		elseif entry.tagname = "sqlsil" then 
		sqlsil = entry.text
		end if
    next
end If

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_eklentiler"
rs.open SQL,baglanti,1,3
rs.addnew
rs("eklenti_adi") = eklenti_adi
rs("eklenti_k") = eklenti_k
rs("eklenti_yazar") = eklenti_yazar
rs("eklenti_mail") = eklenti_mail
rs("eklenti_web") = eklenti_web
rs("eklenti") = aspcode

if not sqlsil = "" then
rs("eklenti_kaldir") = sqlsil
end if
if not sqlcode = "" then
Execute sqlcode
end if
rs.update

next

Response.Write "<br><br><center>Bileşenler Sorunsuz Yüklendi</center><br><br>"
Response.Write "<b>Yüklenen Bileşenler</b><br>* Main Page<br>* Arama Sistemi<br>* Download Sistemi<br>* Galeri Sistemi<br>* Link Sistemi<br>* İletişim Sistemi<br>* Sayfa Sistemi<br>"
Response.Write "<br><br><center><a href=""?islem=adim5"">Devam Et >> </a></center>"
end sub

sub adim5
for zzz="1" to "9"

kurulumadi = "modul"&zzz&".xml"
set xmlDoc = createObject("Microsoft.XMLDOM")
xmlDoc.async = false
xmlDoc.setProperty "ServerHTTPRequest", true
dosya = Server.MapPath(kurulumadi)
xmlDoc.load (dosya)

If (xmlDoc.parseError.errorCode <> 0) then
    Response.Write "XML Hatası: " & xmlDoc.parseError.reason
Else

    set channelNodes = xmlDoc.selectNodes("//item/*")
    for each entry in channelNodes
        if entry.tagName = "moduladi" then
		moduladi = entry.text
        elseif entry.tagname = "aspcode" then 
		aspcode = entry.text
		elseif entry.tagname = "modulyazar" then 
		modulyazar = entry.text
		elseif entry.tagname = "mail" then 
		mail = entry.text
		elseif entry.tagname = "modulizin" then 
		modulizin = entry.text
		elseif entry.tagname = "modulsira" then 
		modulsira = entry.text
		elseif entry.tagname = "modulyer" then 
		modulyer = entry.text
		end if
    next
end If

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from gop_modules"
rs.open SQL,baglanti,1,3
rs.addnew
rs("modul_adi") = moduladi
rs("modul_icerik") = aspcode
rs("modul_yazar") = modulyazar
rs("modul_mail") = mail
rs("modul_izin") = modulizin
rs("modul_sira") = modulsira
rs("modul_yer") = modulyer
rs.update

next

Response.Write "<br><br><center>Modüller Sorunsuz Yüklendi</center><br><br>"
Response.Write "<b>Yüklenen Modüller</b><br>* Ana Menü<br>* Saat<br>* Üye Bilgileri<br>* Son Eklenenler<br>* İstatistik<br>* Kategoriler<br>* Anket<br>* Reklam<br>* Arama<br><br><B><font color=""green"">KURULUM BAŞARIYLA TAMAMLANDI</font></B><br><br><font color=""red"">Lütfen ana dizinden <b>kurulum</b> adlı klasörü güevnliğiniz için siliniz ve <b>Modules</b> ile <b>Upload</b> klasörlerine yazma izni veriniz."
Response.Write "<br><br><center><a href=""../"">Devam Et >> </a></center>"
end sub
%>
</div>
</td>
    </tr>
    <tr>
      <td style="border:solid 1px; border-color:#000000;"><div align="center" class="style1">
      <div align="center">JoomlASP v1.0.3 Kurulum &copy; 2009</div></td>
    </tr>
  </table>