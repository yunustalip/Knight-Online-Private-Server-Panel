<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2007 Web Wiz(TM). All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM 'WEB WIZ'.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN 'WEB WIZ' IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwizguide.com/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwizguide.com
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************


'Global
'---------------------------------------------------------------------------------
Const strTxtForumHelp = "Forum Yard�m�"
Const strTxtChooseAHelpTopic = "Bir yard�m konusu se�in"
Const strTxtLoginAndRegistration = "Forum i�in Kay�t ve Giri�"
Const strTxtUserPreferencesAndForumSettings = "Kullan�c� Tan�mlamalar� ve Forum Ayarlar�"
Const strTxtPostingIssues = "G�nderme-Posta sorunlar�"
Const strTxtMessageFormatting = "Mesaj Bi�imlendirme"
Const strTxtUsergroups = "Kullan�c� gruplar�"
Const strTxtPrivateMessaging = "�zel Haberle�me"

Const strTxtWhyCantILogin = "Neden giremiyorum?"
Const strTxtDoINeedToRegister = "Kay�t olmal� m�y�m?"
Const strTxtLostPasswords = "Parolam� Unuttum"
Const strTxtIRegisteredInThePastButCantLogin = "�nceden kaydolmu�tum, ancak �imdi giremiyorum"

Const strTxtHowDoIChangeMyForumSettings = "Forum ayarlar�m� nas�l de�i�tirebilirim?"
Const strTxtForumTimesAndDates = "Forumun zaman� benim yerel zaman�ma uygun de�il"
Const strTxtWhatDoesMyRankIndicate = "S�r�lamam (rank) neyi g�sterir?"
Const strTxtCanIChangeMyRank = "S�ralamam� (rank) de�i�tirebilir miyim?"
Const strTxtWhatWebBrowserCanIUseForThisForum = "Bu forumda hangi web taray�c�y� kullanmal�y�m?"

Const strTxtHowPostMessageInTheForum = "Forum i�inde nas�l posta g�nderebilirim?"
Const strTxtHowDeletePosts = "Postalar� nas�l silebilirim?"
Const strTxtHowEditPosts = "Postalar� nas�l d�zeltebilirim?"
Const strTxtHowSignaturToMyPost = "Postama nas�l imza ekleyebilirim?"
Const strTxtHowCreatePoll = "Nas�l bir anket yaratabilirim?"
Const strTxtWhyNotViewForum = "Neden forumu g�r�nt�leyemiyorum?"
Const strTxtMyPostIsHiddenOrPendingApproval = "Mesaj�m, �Gizli�, veya �Onay Bekliyor� �eklinde g�r�n�yor"
Const strTxtInternetExplorerWYSIWYGPosting = "Zengin Yaz� Edit�r� (WYSIWYG)"

Const strTxtWhatForumCodes = "Forum kodlar� nelerdir?"
Const strTxtCanIUseHTML = "HTML kullanabilir miyim?"
Const strTxtWhatEmoticons = "Emoticons (Smileys) nedir?"
Const strTxtCanPostImages = "G�r�nt� (images) g�nderebilir miyim?"
Const strTxtWhatClosedTopics = "Kapal� (closed) konular nelerdir?"

Const strTxtWhatForumAdministrators = "Forum Y�neticileri (Forum Administrators) nedir?"
Const strTxtWhatForumModerators = "Forum Ba�kanlar� (Forum Moderators) nedir?"
Const strTxtWhatUsergroups = "Kullan�c� Gruplar� (Usergroups) nedir?"

Const strTxtWhatIsPrivateMessaging = "�zel Mesaj Sistemi nedir?"
Const strTxtIPrivateMessages = "�zel mesajlar�m� g�nderemiyorum"
Const strTxtIPrivateMessagesToSomeUsers = "�zel mesajlar�m� sadece baz� kullan�c�lara g�nderemiyorum"
Const strTxtHowCanPreventSendingPrivateMessages = "Baz�lar�n�n bana �zel mesaj g�ndermesini nas�l engelleyebilirim?"

Const strTxtRSSFeeds = "RSS Linkleri"
Const strTxtHowDoISubscribeToRSSFeeds = "Forum RSS Linklerine Nas�l Abone Olabilirim?"

Const strTxtCalendarSystem = "Takvim Sistemi"
Const strTxtWhatIsTheCalendarSystem = "Takvim Sistemi Nedir?"
Const strTxtHowDoICreateCalendarEvent = "Nas�l takvim olay� giri�i yapabilirim?"

Const strTxtAbout = "Hakk�nda"
Const strTxtWhatSoftwareIsUsedForThisForum = "Bu forumda hangi yaz�l�m kullan�l�yor?"


Const strTxtFAQ1 = "Foruma girebilmek i�in kay�t yapt�r�rken yazd���n�z kullan�c� ad� ve parolas�n� girmeniz gerekmektedir.Halen kay�tl� �ye de�ilseniz foruma girebilmek i�in �ncelikle kay�tl� �ye olmal�s�n�z. Kay�tl� �ye olman�za ve do�ru kullan�c� ad� ile parolay� yazman�za ra�men halen foruma giremiyorsan�z, �ncelikle web taray�c�n�zda �erezlerin (cookies) kullan�labilir olup olmad���n� kontrol edin. Halen sorun ya��yorsan�z web taray�c�n�za bu siteyi g�venilir site olarak tan�tman�z gerekebilir. E�er foruma giri�iniz yasaklanm�� (banned) ise yine foruma giremeyece�inizden her halde forum y�neticiniz ile g�r���n�z."
Const strTxtFAQ2 = "Foruma g�nderme yapabilmek i�in kay�t olman�z gerekmeyebilir, ��nk� kay�tl� olmasa dahi kullan�c�lar�n foruma g�nderme yapma yetkisini verme ya da yasaklama forum y�neticisinin elindedir. Ancak, foruma kay�t olmakla ek �zelliklerden ve olanaklardan yararlanabilece�inizi dikkate al�n�z�. Kay�t olmak sadece bir ka� dakikan�z� alaca��ndan, kay�t olman�z� �neriyoruz."
Const strTxtFAQ3 = "Parolan�z� unuttu�unuzda tela�lanman�za gerek yok. �nceki parolalar hat�rlanamaz ise de, yerine yenisi verilebilir. Unuttu�unuzun yerine yeni parola almak i�in giri� (login) d��mesine bas�n ve giri� sayfas�n�n alt�nda g�rece�iniz parolam� unuttum (lost password) sayfas� linkini t�klad���n�zda y�nlendirildi�iniz sayfadan, yeni parolan�n e posta ile g�nderilmesini isteyebilirsiniz. E�er belirtilen se�enek kullan�labilir de�il ya da, yeni parolan�n g�nderilebilece�i bir e posta adresi �nceden profilinizde belirtilmemi� ise, forum y�netecisi ya da forum ba�kan� ile temasa ge�erek parolan�z� sizin i�in de�i�tirmesini isteyebilirsiniz."
Const strTxtFAQ4 = "Bir s�re i�in g�nderme yapmam�� ya da hi� g�nderme yapmam�� olabirsiniz. Forum y�neticilerinin veritaban�ndan uzun belirli s�re foruma kat�lmayan kullan�c�lara, veritaban�n� bo�alt�p rahatlatmak i�in silmeleri al���lagelmi� bir uygulamad�r."
Const strTxtFAQ5 = "Forum ayarlar�n�z�, profil bilgilerinizi, kay�t detaylar�n�z� vs., <a href=""member_control_panel.asp"" target=""_self"">�ye Kontrol Paneli</a>, den foruma bir kez girmi� iseniz de�i�tirebilirsiniz. Bu merkez men�den bir �ok g�r�n�m� kontrol edebilir ve �ye olanaklar�na eri�ebilirsiniz."
Const strTxtFAQ6 = "Forumlarda kullan�lan zaman, sunucunun tarih ve saati oldu�undan, e�er sunucu ba�ka bir �lkede bulunuyor ise, tarih ve saat o �lkenin yerel ayarlar�d�r. Tarih ve saati kendi yerel ayarlar�n�za de�i�tirmek i�in ('Forum Preferences') Forum Tercihlerinizi <a href=""member_control_panel.asp"" target=""_self"">�ye Kontrol Paneli</a> den de�i�tirip, sunucu ile yerel ayarlar�n�z aras�ndaki fark� yaz�p ayarlar�n�z� yap�n�z. Forumlar standart ve g�n ����� tasarruf zamanlamas�n� ayarlayabilmek i�in tasarlanmad���ndan, g�n ����� tasarrufu i�in yap�lan ayarlama zamanlar�nda yeniden zaman ayarlar�n� yapman�z gerekebilir."
Const strTxtFAQ7 = "Forumdaki s�ralama (ranks) sizin hangi kullan�c� grubunun �yesi oldu�unuzu ve kullan�c�lar� g�sterir. �rne�in, forum ba�kanlar� ve y�neticilerinin farkl� ve �zel bir s�ralamas� olabilir. Forumun ayarlar�na ve hangi s�ralamada oldu�unuza ba�l� olarak, forumun de�i�ik �zelliklerini kullanabilirsiniz."
Const strTxtFAQ8 = "Yapamaman�z normaldir. Ancak, forum y�neticisi y�kselme sistemini (ladder system) kullanarak s�ralamay� yapm��sa, mesajlar�n�z�n say�s�na g�re bir �st gruba otomatik olarak ge�ersiniz."
Const strTxtFAQ9 = "Forumlarda mesaj g�nderebilmek i�in, forum ya da konu ekran�nda ilgili d��meyi t�klat�n. Forum y�neticisinin forum ayarlar�na ba�l� olarak, bir mesaj g�ndermeden �nce foruma girmeniz gerekebilir. Konu ekran�n�n alt�nda kullanabilece�iniz kolayl�klar listelenmi�tir."
Const strTxtFAQ10 = "Forumun g�ndermelerinizi silebilmeniz i�in uygun ayarlara sahip olmas� kayd�yla forum ba�kan� ya da forum y�neticisi olmad���n�z ve s�rece, sadece kendi g�ndermelerinizi silebilirsiniz. E�er biri sizin g�ndermenizi cevaplam�� ise, art�k kendi g�ndermenizi de silemezsiniz."
Const strTxtFAQ11 = "Forumun g�ndermelerinizi d�zenleyebilmeniz i�in uygun ayarlara sahip olmas� kayd�yla, forum ba�kan� ya da forum y�neticisi olmad���n�z s�rece g�ndermelerinizi d�zenleyemezsiniz. Forum ayarlar�na ba�l� olarak g�ndermelerinizi d�zeltti�inizde, d�zelten kullan�c� ad�, d�zeltmenin tarih ve saati g�nderimin alt�nda g�r�necektir."
Const strTxtFAQ12 = "Forum y�neticisi forumda imza kullan�m�na izin vermi� ise, g�nderiminizin alt�na imzan�z� ekleyebilirsiniz. �mza ekleyebilmek i�in �ncelikle ('Profile Information') Ki�isel Bilgilerinizde <a href=""member_control_panel.asp"" target=""_self"">�ye Kontrol Paneli</a>, ni kullanarak imzan�z� yaratm�� olman�z gerekmektedir. Bir kez bunu yapt�ktan sonra G�nderme formunun alt�ndaki ('Show Signature') �mzay� G�ster kutusunu i�aretleyerek imzan�z� g�nderebilirsiniz."
Const strTxtFAQ13 = "Bir forumda anket yaratmak i�in yeterli haklar�n�z varsa, forum ve konular ekran�n�n tepesinde ('New Poll') Yeni Anket d��mesini g�receksiniz. Anket yarat�rken, anket sorusunu ve en az iki anket se�ene�i girmeniz gerekmektedir. Ayr�ca, kat�lanlar�n sadece bir defa ya da bir defadan fazla oy kullanabileceklerini se�ebilirsiniz."
Const strTxtFAQ14 = "Baz� forumlar, sadece belirli kullan�c�lar�n ya da belirli kullan�calar�n olu�turdu�u gruplar�n eri�imi i�in ayarlanm��t�r. Bu t�r bir forumu g�rmek, okumak vs. i�in �ncelikle, forum ba�kan� ya da y�neticisinin sa�layabilece�i izinlere ihtiyac�n�z olacakt�r."
Const strTxtFAQ28 = "Onay gerektiren bir foruma veya konuya mesaj yazm��s�n�z. Mesaj�n�z� herkesin g�rebilmesi i�in y�netici veya moderat�rler taraf�ndan onaylanmas� gerekir."
Const strTxtFAQ15 = "Internet Explorer 5+ (sadece windows), Netscape 7.1, Mozilla 1.3+, Mozilla Firebird 0.6.1+, kullan�yorsan�z ve forum y�netici zengin yaz� edit�r�n� a�m��sa zengin yaz� edit�r�n� kullanarak mesaj yazabilirsiniz. Zengin yaz� edit�r�nde sorun ya��yorsan�z ayarlar b�l�m�nden kapatabilirsiniz."
Const strTxtFAQ16 = "Forum kodlar�, forumda g�nderdi�iniz mesajlar� bi�imlendirmenize olanak verir.Forum kodlar� HTML ye �ok benzer. Sadece tag'lar k��eli parantez i�indedir. �rne�in, &lt; and &gt; yerine [ and ] gibi. Ayr�ca, mesaj g�nderirken forum kodlar�n� �al��t�rmayabilirsiniz. <a href=""javascript:winOpener('forum_codes.asp','codes',1,1,550,400)"">Buray� t�klayarak kullanabilece�iniz forum kodlar�n� g�rebilirsiniz</a>."
Const strTxtFAQ17 = "Mesajlar�n�zda HTML kullan�lamaz. K�t� niyetle g�nderilebilecek HTML kodlar�n�n forumun g�r�nt�s�n� bozabilece�inden ve hatta mesaj kullan�c� taraf�ndan a��ld���nda web taray�c�s� �al��amaz duruma gelebilece�inden g�venlik nedeniyle kullan�lmamaktar."
Const strTxtFAQ18 = "Emoticons or Smileys'ler duygular� ifade etmek ya da g�stermek i�in kullan�labilecek k���k grafik g�r�nt�lerdir. E�er forum y�neticisi izin vermi� ise Emoticons'u forumlarda mesaj g�nderirken mesaj formunun yan�nda g�rebilirsiniz. G�nderiminize emoticon eklemek i�in, g�ndermek istedi�iniz emoticon �zerine t�klay�n�z."
Const strTxtFAQ19 = "E�er forum y�neticiniz y�klemeye izin vermi�se, bilgisayar�n�zdaki bir g�r�nt� postan�za y�klenebilir (upload) ve eklenebilir. Ancak, e�er g�r�nt�n�n (image) y�klenebilmesi m�mk�n de�il ise, g�r�nt�n�n bulundu�u sunucuya link vermeniz gerekecektir. �rne�in, http://www.mysite.com/my-picture.jpg."
Const strTxtFAQ20 = "Kapal� Konular (Closed Topics), forum ba�kanlar� ya da y�neticileri taraf�ndan bu �ekilde ayarlanm�� olanlard�r. E�er bir kapat�lm�� ise, art�k bu konuda g�nderme yapamaz, g�ndermeyi cevaplayamaz ve bir anketi oylayamazs�n�z."
Const strTxtFAQ21 = "Forum Y�neticileri (Forums Administrators) forumlarda en y�ksek kontrol yetkisine sahip ki�ilerdir. Forumlarda �zellikleri kapat�p a�ma, kullan�c�lar� yasaklama, kullan�c�lar� silme, g�ndermeleri d�zeltme ve silme, kullan�c� gruplar� olu�turma gibi imkanlar� bulunmaktad�r."
Const strTxtFAQ22 = "Forum Ba�kanlar� (Moderators) forumun gidi�ine g�nl�k olarak bakan tek ya da grup halinde kullan�c�lard�r. Sorumlu olduklar� forumda, konu ve g�ndermeleri d�zeltme, silme, yer de�i�tirme, kapatma, kapat�lm��� a�ma haklar�na sahiptirler. Bunlar genellikle ki�ilerin sald�rgan ve k�f�rl� ifadelerini engellemek i�in bulunurlar."
Const strTxtFAQ23 = "Kullan�c� gruplar�, grup kullan�c�lar�n bir yoludur. Her kullan�c� bir kullan�c� grubunun �yesidir ve her gruba forumda kullanabilece�i okuma, g�rme, g�nderme, anket yaratma gibi kendisine �zg� haklar tan�nm�� olabilir."
Const strTxtFAQ24 = "Bunun bir ka� nedeni olabilir. Foruma girmemi� olabilirsiniz, kay�tl� olmayabilirsiniz veya forum y�neticileri �zel Mesaj G�nderme sistemini (Private Messaging system) iptal etmi� olabilir."
Const strTxtFAQ25 = "Bu, �zel mesaj g�ndermek istedi�iniz ki�ininin sizden gelecek �zel mesajlar� bloke etmesinden kaynaklan�yon olabilir. Hal b�yle ise, �zel mesaj g�nderdi�inizde bu konuda sizi uyaran bir mesaj alm�� olman�z gerekir"
Const strTxtFAQ26 = "E�er bir kullan�c�dan istenmeyen �zel mesajlar al�yorsan�z, size �zel mesaj g�nderilmesini bloke edebilirsiniz. Blokaj� yapmak i�in �zel Mesaj Sisteminde (Private Messaging system) Arkada� listenize (buddy list) gidin. Kullan�c�y� Arkada� (buddy) olarak yaz�n, fakat drop down listeden ('Not to message you') Mesaj g�ndermesini se�in. B�ylece bu kullan�c�n�n size �zel mesaj g�ndermesini engellemi� olacaks�n�z."
Const strTxtFAQ27 = "E�er forum y�neticisi �zel mesaj sistemini a�m��sa, di�er �yelere mesaj g�nderebilir ve alabilirsiniz."
Const strTxtFAQ29 = "RSS, sitelerin g�ncel i�eriklerini payla�malar�n� sa�layan bir teknolojidir. E�er forum y�neticisi bu �zelli�i etkinle�tirmi� ise foruma eklenen son mesajlardan ve takvimdeki �nemli olaylardan an�nda haberdar olabilirsiniz."
Const strTxtFAQ30 = "Forumdaki RSS deste�inden yararlanman�z�n bir �ok y�ntemi vard�r, �rne�in; RSS Deste�i olan bir Web Browser kullanmak (Firefox, IE7, Safari, Opera), RSS Haber Okuyucu kullanmak, veya RSS deste�i olan Mozilla Thunderbird gibi bir e-posta program� kullanmak."
Const strTxtFAQ31 = "Takvime girilmi� olan mesajlar� ve �yelerin do�um g�nlerini g�nl�k, haftal�k ve ayl�k olarak g�rebilirsiniz."
Const strTxtFAQ32 = "Forum y�neticisi sizin i�inde bulundu�unuz gruba veya size �zel takvime giri� yapma yetkisi vermi� ise takvime olay giri�i yapabilirsiniz. E�er izniniz varsa mesaj yazarken tarih se�ebilece�iniz bir alan gelecektir."
Const strTxtFAQ33 = "Web wiz forum kullan�lmaktad�r. Web Wiz Forums forum Microsoft�s Active Server Pages (ASP) taban�yla kodlanm��t�r ve Windows web serverlarda �al���r. Kendi web siteniz i�in web wiz forum indirmek i�in t�klay�n <a href=""http://www.webwizforums.com"">www.webwizforums.com</a>."
Const strTxtFAQ34 = "Windows XP veya Apple MAC OS X platformunda Internet Explorer, Mozilla, Firefox, Safari, Netscape, Opera, ve benzeri taray�c�lar� kullanabilirsiniz. Ama biz t�m platformlar i�in <a href=""http://getfirefox.com"" target=""_blank"">Firefox</a> u �neriyoruz."

%>