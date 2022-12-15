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
Baglanti.open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="&mysql_server&"; UID="&mysql_user&"; pwd="&mysql_pass&";db="&mysql_db&"; option = 999999"

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

Baglanti.execute "DROP TABLE IF EXISTS `gop_adminmenua`;"
sql = "CREATE TABLE `gop_adminmenua` ( "
sql = sql & "`amid` INT(11) NOT NULL auto_increment,"
sql = sql & "`amadi` varchar(55) default NULL, "
sql = sql & "`amar` varchar(255) default NULL, "
sql = sql & "`amka` varchar(50) default NULL, "
sql = sql & "`amord` INT(11) default NULL, "
sql = sql & "PRIMARY KEY  (amid)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Admin Menu Tablosu Oluşturuldu<br>"



Baglanti.execute "DROP TABLE IF EXISTS `gop_adminmenub`;"
sql = "CREATE TABLE `gop_adminmenub` ("
sql = sql & "`ambid` int(11) NOT NULL auto_increment,"
sql = sql & "`ambadi` varchar(55) default NULL,"
sql = sql & "`amblink` varchar(255) default NULL,"
sql = sql & "`amka` varchar(50) default NULL,"
sql = sql & "PRIMARY KEY  (`ambid`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Admin Alt Menu Tablosu Oluşturuldu<br>"


Baglanti.execute "DROP TABLE IF EXISTS `gop_anketip`;"
Baglanti.execute "DROP TABLE IF EXISTS `ip_block`;"
sql = "CREATE TABLE `gop_anketip` ("
sql = sql & "`no` int(10) NOT NULL auto_increment,"
sql = sql & "`poll_id_ip` int(10) NOT NULL default '0',"
sql = sql & "`ip` varchar(25) default '''255.255.255.255''',"
sql = sql & "PRIMARY KEY  (`no`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Anket IP Tablosu Değiştirildi<br>"



Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN siteadres varchar(255) default NULL;"
Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN kresim int(11) default NULL;"
Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN google int(11) default NULL;"
Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN google_code varchar(255) default NULL;"
Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN yenileme int(3) NOT NULL default '240';"
Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN eklenti_k varchar(25) NOT NULL default 'main_page';"
Baglanti.execute "ALTER TABLE gop_ayarlar ADD COLUMN dilayari int(11) NOT NULL default '1';"
Response.Write "* Ayarlar Tablosu Güncelleştirildi<br>"


Baglanti.execute "DROP TABLE IF EXISTS `gop_eklentiler`;"
sql = "CREATE TABLE `gop_eklentiler` ("
sql = sql & "`id` int(11) NOT NULL auto_increment,"
sql = sql & "`eklenti_adi` varchar(55) default NULL,"
sql = sql & "`eklenti` mediumtext,"
sql = sql & "`eklenti_yazar` varchar(55) default NULL,"
sql = sql & "`eklenti_mail` varchar(255) default NULL,"
sql = sql & "`eklenti_k` varchar(55) default NULL,"
sql = sql & "`eklenti_web` mediumtext,"
sql = sql & "`eklenti_kaldir` mediumtext,"
sql = sql & "PRIMARY KEY  (`id`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Bileşenler Tablosu Oluşturuldu<br>"


Baglanti.execute "DROP TABLE IF EXISTS `gop_iletisim`;"
sql = "CREATE TABLE `gop_iletisim` ("
sql = sql & "`id` int(11) NOT NULL auto_increment,"
sql = sql & "`adi` varchar(50) default NULL,"
sql = sql & "`mail` varchar(255) default NULL,"
sql = sql & "`yas` int(2) default NULL,"
sql = sql & "`tel` varchar(25) default NULL,"
sql = sql & "`mesaj` mediumtext,"
sql = sql & "`tarih` datetime NOT NULL default '2008-01-01 12:00:00',"
sql = sql & "PRIMARY KEY  (`id`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* İletişim Tablosu Oluşturuldu<br>"



Baglanti.execute "DROP TABLE IF EXISTS `gop_language`;"
sql = "CREATE TABLE `gop_language` ("
sql = sql & "`lang_id` int(11) NOT NULL auto_increment,"
sql = sql & "`language` mediumtext,"
sql = sql & "`lang_adi` varchar(25) default NULL,"
sql = sql & "`lang_yazar` varchar(50) default NULL,"
sql = sql & "`lang_mail` varchar(255) default NULL,"
sql = sql & "PRIMARY KEY  (`lang_id`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Dil Tablosu Oluşturuldu<br>"


Baglanti.execute "DROP TABLE IF EXISTS `gop_mesajlar`;"
sql = "CREATE TABLE `gop_mesajlar` ("
sql = sql & "`mesaj_id` int(11) NOT NULL auto_increment,"
sql = sql & "`yollayan` int(11) default NULL,"
sql = sql & "`alici` int(11) default NULL,"
sql = sql & "`mesaj_baslik` varchar(50) default NULL,"
sql = sql & "`mesaj_icerik` mediumtext,"
sql = sql & "`mesaj_tarih` datetime default NULL,"
sql = sql & "`mesaj_sil` int(1) NOT NULL default '0',"
sql = sql & "`mesaj_okundu` int(1) NOT NULL default '0',"
sql = sql & "`mesaj_gsil` int(1) NOT NULL default '0',"
sql = sql & "PRIMARY KEY  (`mesaj_id`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Mesajlar Tablosu Oluşturuldu<br>"



Baglanti.execute "ALTER TABLE gop_sayac ADD COLUMN sayac_tarih date NOT NULL default '0000-00-00';"
Response.Write "* İstatistik Tablosu Güncelleştirildi<br>"



Baglanti.execute "DROP TABLE IF EXISTS `gop_sayacayar`;"
sql = "CREATE TABLE `gop_sayacayar` ("
sql = sql & "`id` int(11) NOT NULL auto_increment,"
sql = sql & "`btekil` int(11) NOT NULL default '0',"
sql = sql & "`bcogul` int(11) NOT NULL default '0',"
sql = sql & "`toplamc` int(11) NOT NULL default '0',"
sql = sql & "`toplamt` int(11) NOT NULL default '0',"
sql = sql & "`dtekil` int(11) NOT NULL default '0',"
sql = sql & "`dcogul` int(11) NOT NULL default '0',"
sql = sql & "`aktifuye` int(11) NOT NULL default '0',"
sql = sql & "`okunma` int(11) NOT NULL default '0',"
sql = sql & "`ip` int(11) NOT NULL default '0',"
sql = sql & "`online` int(11) NOT NULL default '0',"
sql = sql & "`veri` int(11) NOT NULL default '0',"
sql = sql & "`sonuye` int(11) NOT NULL default '0',"
sql = sql & "PRIMARY KEY  (`id`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* İstatistik Ayarlar Tablosu Oluşturuldu<br>"



Baglanti.execute "DROP TABLE IF EXISTS `gop_sayfa`;"
sql = "CREATE TABLE `gop_sayfa` ("
sql = sql & "`sayfaid` int(11) NOT NULL auto_increment,"
sql = sql & "`sayfa_baslik` varchar(255) default NULL,"
sql = sql & "`sayfa_icerik` mediumtext,"
sql = sql & "`sayfa_hit` int(11) NOT NULL default '0',"
sql = sql & "`sayfa_tarih` datetime NOT NULL default '0000-00-00 00:00:00',"
sql = sql & "PRIMARY KEY  (`sayfaid`)"
sql = sql & ") ENGINE=InnoDB;"
baglanti.Execute sql
Response.Write "* Sayfalar Tablosu Oluşturuldu<br>"


Baglanti.execute "ALTER TABLE gop_uyeler ADD COLUMN uye_dil int(11) NOT NULL default '1';"
Baglanti.execute "ALTER TABLE gop_uyeler ADD COLUMN uye_ip varchar(25) NOT NULL default '255.255.255.255';"
Response.Write "* Üyeler Tablosu Güncelleştirildi.<br>"



Baglanti.execute "ALTER TABLE gop_veriler ADD COLUMN vetiket mediumtext NOT NULL;"
Response.Write "* Veriler Tablosu Güncelleştirildi.<br>"

Response.Write "<br><br><center><a href=""?islem=adim2"">Devam Et >> </a></center></td></tr></table>"
end sub

sub adim2

Response.Write "<br><br>Veriler giriliyor.<br><br><table><tr><td align=""left"">"
Response.Write "* Admin ana menüsü işleniyor.<br>"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (1,'Response.Write general_settings','genel_ayarlar.png','genel_ayarlar',1);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (2,'Response.Write menu_settings','menu_islemleri.png','menu_islemleri',3);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (3,'Response.Write category_settings','kategori_islemleri.png','kategori_islemleri',4);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (4,'Response.Write data_settings','veri_islemleri.png','veri_islemleri',5);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (5,'Response.Write addition_settings','modul_islemleri.png','eklenti_islemleri',6);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (6,'Response.Write members_settings','uye_islemleri.png','uye_islemleri',7);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (7,'Response.Write link_settings','link_islemleri.png','link_islemleri',8);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (8,'Response.Write download_settings','download_islemleri.png','download_islemleri',9);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (9,'Response.Write gallery_settings','galeri_islemleri.png','galeri_islemleri',10);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (10,'Response.Write poll_settings','anket_yonetimi.png','anket_islemleri',11);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (11,'Response.Write advertising_settings','reklam_yonetimi.png','reklam_islemleri',12);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (12,'Response.Write communication_settings','iletisim_islemleri.png','iletisim_islemleri',13);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (13,'Response.Write language_settings','dil_islemleri.png','dil_islemleri',2);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (14,'Response.Write page_settings','veri_islemleri.png','sayfa_islemleri',5);"
Baglanti.execute "INSERT INTO `gop_adminmenua` VALUES (15,'Response.Write statistics ','sayac_ayarlari.asp','sayac_islem',14);"
Response.Write "<b>* İşlendi.</b><br>"


Response.Write "* Admin menüsü işleniyor.<br>"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (1,'Response.Write site_settings','site_ayarlari.asp','genel_ayarlar');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (2,'Response.Write menu','menuler.asp','menu_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (3,'Response.Write add','menuler.asp?islem=menu_ekle','menu_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (4,'Response.Write category','kategoriler.asp','kategori_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (5,'Response.Write add','kategoriler.asp?act=ekle','kategori_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (6,'Response.Write sub_category','altkategoriler.asp','kategori_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (7,'Response.Write add','altkategoriler.asp?islem=yeni','kategori_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (8,'Response.Write data','veriler.asp','veri_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (9,'Response.Write add','veriler.asp?islem=ekle','veri_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (10,'Response.Write comment','yorumlar.asp','veri_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (11,'Response.Write review_approve','yorumlar.asp?islem=bekleyen','veri_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (12,'Response.Write modules','moduller.asp','eklenti_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (13,'Response.Write add_manuel_module','moduller.asp?islem=modul_manuel_ekle','eklenti_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (14,'Response.Write members','uyeler.asp','uye_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (15,'Response.Write add','uyeler.asp?islem=ekle','uye_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (16,'Response.Write links','linkler.asp','link_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (17,'Response.Write add','linkler.asp?act=ekle','link_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (18,'Response.Write approve','linkler.asp?act=onay','link_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (19,'Response.Write category','downloadlar.asp?islem=kategoriler','download_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (20,'Response.Write downloads','downloadlar.asp','download_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (21,'Response.Write add','downloadlar.asp?islem=ekle','download_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (22,'Response.Write approve','downloadlar.asp?islem=onay','download_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (23,'Response.Write category','galeri.asp','galeri_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (24,'Response.Write pictures','galeri.asp?islem=resimler','galeri_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (25,'Response.Write polls','anket.asp','anket_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (26,'Response.Write add','anket.asp?sub=addnew','anket_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (27,'Response.Write advertisement','reklamlar.asp','reklam_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (28,'Response.Write inbox','iletisim.asp','iletisim_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (29,'Response.Write add_component','bilesenler.asp','eklenti_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (30,'Response.Write language','diller.asp','dil_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (31,'Response.Write database_backup','db_yedekle.asp','genel_ayarlar');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (32,'Response.Write pages','sayfalar.asp','sayfa_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (33,'Response.Write add_page','sayfalar.asp?islem=ekle','sayfa_islemleri');"
Baglanti.execute "INSERT INTO `gop_adminmenub` VALUES (34,'Response.Write statistics ','sayac_ayarlari.asp','sayac_islem');"
Response.Write "<b>* İşlendi.</b><br>"


Response.Write "* Mesajlar işleniyor<br>."
Baglanti.execute "INSERT INTO `gop_mesajlar` VALUES (1,1,1,'Hoşgeldiniz...','JoomlASP güncellemesi başarıyla tamamlamış bulunmaktasınız. Kullanım ve Destek için JoomlASP Resmi sitesi olan www.joomlasp.com adresinizi ziyaret edebilirsiniz.','2009-02-06 16:52:29',0,0,0);"
Response.Write "<b>* İşlendi.</b><br>"


Response.Write "* İstatistik Ayarları işleniyor.<br>"
Baglanti.execute "INSERT INTO `gop_sayacayar` VALUES (1,1,1,1,1,1,1,1,1,1,1,1,1);"
Response.Write "<b>* İşlendi.</b><br>"

Response.Write "* Sayfalar işleniyor.<br>"
Baglanti.execute "INSERT INTO `gop_sayfa` VALUES (1,'JoomlASP','JoomlASP, Asp (Active Server Page) dilinde yazilmis olan ve birebir Joomla özeliklerini kapsayan ve gelisime açik bir cms sistemidir. Kisisel web sitelerinden, özel firma sitelerine kadar bir siteyi rahatlikla yapabileceginiz, kolay, modüler, mysql db sistemi hizli bir sekilde islem yapabilen bir cms sistemidir.<br><br>\r\n \r\nSistemin yazilis amaci Php dilinin bir çok kisiye basit gelmesine karsin Asp dili ile internet programciligina baslayan kisilerde Php dilinden uzak kalma ve Php kodlarini çözememe gibi durumlar için hazirlanmistir. Asp dilini uzan yillardir kullanan JoomlASP yapimcisi Hasan Emre ASKER, Php dilinden anlamayan ve karmasik yapiya sahip oldugunu düsünen kisiler için kullanimi Joomlaya göre daha kolay olan bu CMS sistemini meydana getirmistir.<br><br>\r\n\r\nSistemimiz gelistirilmeye açik bir yapiya sahip oldugundan ve Joomla temalari ile birebir uyumluluk sergilediginden begendiginiz Joomla temasini basit bir kaç editleme ile JoomlASP\'ye uyarlayabilirsiniz. Eger sizler bunu yapacak vaktimiz yok ben su temanin JoomlASP versiyonunu istiyorum derseniz, site editörlerimiz tarafindan istediginiz bir Joomla temasi JoomlASP sistemine çevrilip temalar bölümünde yayinlanacaktir.<br><br> \r\n\r\nYine sistem için modül yazimini kolaylastirmak için site yapimizi basitlestirmis ve asp dilinden çok az anlayan birinin bile kendine göre küçük modüller yazabilecegine inancimiz tamdir.<br><br> \r\n\r\nSitemiz haricinde yayinlanan JoomlASP bilesen ve Modüllerinden lütfen uzak durunuz. Yazilmis olan Modül, Bilesen ve Temalar sitemiz editörleri tarafindan incelenip açiklarinin var olup olmadigi kontrol edildikten sonra yayina hazir halde sizlere sunulacaktir. Bu yüzden lütfen disardan temin ettiginiz Bilesen, Modül ve Temalari kullanmayiniz. ',546,'2008-10-25 10:00:00');"
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
Response.Write "<b>Yüklenen Modüller</b><br>* Ana Menü<br>* Saat<br>* Üye Bilgileri<br>* Son Eklenenler<br>* İstatistik<br>* Kategoriler<br>* Anket<br>* Reklam<br>* Arama<br><br><B><font color=""green"">GÜNCELLEME BAŞARIYLA TAMAMLANDI</font></B><br><br><font color=""red"">Lütfen ana dizinden <b>kurulum</b> adlı klasörü güvenliğiniz için siliniz ve <b>Modules</b> ile <b>Upload</b> klasörlerine yazma izni veriniz."
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