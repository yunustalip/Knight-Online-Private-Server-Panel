<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2011 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwiz.co.uk/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************


'Global
'---------------------------------------------------------------------------------
Const strTxtWelcome = "Hoþ Geldiniz"
Const strTxtAllForums = "Tüm Forumlar"
Const strTxtTopics = "Konular"
Const strTxtPosts = "Mesajlar"
Const strTxtLastPost = "Son Mesaj"
Const strTxtPostPreview = "Mesaj ön izleme"
Const strTxtAt = "saat"
Const strTxtBy = "yazan:"
Const strTxtOn = ""
Const strTxtProfile = "Kullanýcý profili"
Const strTxtSearch = "Arama"
Const strTxtQuote = "Alýntý"
Const strTxtVisit = "Ziyaret"
Const strTxtView = "Gösterim"
Const strTxtHome = "Ana"
Const strTxtHomepage = "Web Siteniz"
Const strTxtEdit = "Düzenle"
Const strTxtDelete = "Sil"
Const strTxtEditProfile = "Profili Düzenle"
Const strTxtLogOff = "Çýkýþ"
Const strTxtRegister = "Foruma Kayýt Olun"
Const strTxtLogin = "Giriþ"
Const strTxtMembersList = "Forum Üyelerini Göster"
Const strTxtForumLocked = "Forum Kilitli"
Const strTxtSearchTheForum = "Forum Aramasý"
Const strTxtPostReply = "Cevap Yaz"
Const strTxtNewTopic = "Yeni Konu"
Const strTxtNoForums = "Gösterilecek hiç forum yok"
Const strTxtReturnToDiscussionForum = "Foruma Geri Dön"
Const strTxtMustBeRegistered = "Bu özelliði kullanabilmek için foruma kayýt olmalýsýnýz."
Const strTxtClearForm = "Formu Temizle"
Const strTxtYes = "Evet"
Const strTxtNo = "Hayýr"
Const strTxtForumLockedByAdmim = "Giriþ engellendi.<br />Bu Forum Yönetici tarafýndan kilitlenmiþtir."
Const strTxtRequiredFields = "Ýþaretli alanlar zorunludur"

Const strTxtForumJump = "Foruma Atla"
Const strTxtSelectForum = "Forum Seçiniz"

Const strTxtNoMessageError = "Mesaj Kutusu \t\t- Mesaj alaný boþ olamaz."
Const strTxtErrorDisplayLine = "_______________________________________________________________"
Const strTxtErrorDisplayLine1 = "Formdaki hatalardan dolayý iþleminiz tamamlanamadý."
Const strTxtErrorDisplayLine2 = "Lütfen hatalarý giderdikten sonra tekrar deneyiniz."
Const strTxtErrorDisplayLine3 = "Aþaðýdaki alan(lar) düzeltilmelidir: -"



'default.asp
'---------------------------------------------------------------------------------
Const strTxtCookies = "Forumu kullanabilmek için Çerezler'in ve JavaScript'in tarayýcýnýzdan açýk olmasý gerekmektedir."
Const strTxtForum = "Forum"
Const strTxtLatestForumPosts = "Son Forum Mesajlarý"
Const strTxtForumStatistics = "Forum Ýstatistikleri"
Const strTxtNoForumPostMade = "Forumda hiç mesaj bulunmamaktadýr"
Const strTxtThereAre = "Toplam"
Const strTxtPostsIn = "Mesaj,"
Const strTxtTopicsIn = "Konu," 'Konularýn içinde yazmak mantýksýz
Const strTxtLastPostBy = "Last Post by"
Const strTxtForumMembers = "Forum Üyemiz vardýr"
Const strTxtTheNewestForumMember = "En yeni forum üyemiz"


'forum_topics.asp
'---------------------------------------------------------------------------------
Const strTxtThreadStarter = "Konuyu Açanlar"
Const strTxtReplies = "Cevaplar"
Const strTxtViews = "Okunma"
Const strTxtDeleteTopicAlert = "Bu konuyu silmek istediðinizden emin misiniz?"
Const strTxtDeleteTopic = "Konuyu Sil"
Const strTxtNextTopic = "Sonraki Konu"
Const strTxtLastTopic = "Son Konu"
Const strTxtShowTopics = "Konularý göster"
Const strTxtNoTopicsToDisplay = "Bu forumun içinde gösterilecek hiç mesaj yok"

Const strTxtAll = "Hepsi"
Const strTxtLastWeek = "Son Hafta"
Const strTxtLastTwoWeeks = "Son iki Hafta"
Const strTxtLastMonth = "Son Ay"
Const strTxtLastTwoMonths = "Son Ýki Ay"
Const strTxtLastSixMonths = "Son Altý Ay"
Const strTxtLastYear = "Son Yýl"


'forum_posts.asp
'---------------------------------------------------------------------------------
Const strTxtLocation = "Konum"
Const strTxtJoined = "Kayýt tarihi"
Const strTxtForumAdministrator = "Forum Yöneticisi"
Const strTxtForumModerator = "Forum Operatörü"
Const strTxtDeletePostAlert = "Bu mesajý silmek istediðinizden emin misiniz?"
Const strTxtEditPost = "Mesajý Düzenle"
Const strTxtDeletePost = "Mesajý Sil"
Const strTxtSearchForPosts = "Mesajlarda ara"
Const strTxtSubjectFolder = "Baþlýk"
Const strTxtPrintVersion = "Yazdýr"
Const strTxtEmailTopic = "Email gönder"
Const strTxtSorryNoReply = "Eriþim engellendi."
Const strTxtThisForumIsLocked = "Bu forum yönetici tarafýndan kilitlenmiþtir."
Const strTxtPostAReplyRegister = "Eðer bu konuya mesaj yollamak istiyorsanýz öncelikle"
Const strTxtNeedToRegister = "Eðer foruma kayýt olmamýþsanýz öncelikle kayýt olmalýsýnýz"
Const strTxtSmRegister = "kayýt"
Const strTxtNoThreads = "Bu konuya ait hiçbir mesaj bulunamadý"
Const strTxtNotGiven = "Girilmedi"


'search_form.asp
'---------------------------------------------------------------------------------
Const strTxtSearchFormError = "Arama\t- Lütfen aranacak kelimeyi yazýn"


'search.asp
'---------------------------------------------------------------------------------
Const strTxtSearchResults = "Arama Sonuçlarý"
Const strTxtHasFound = "bulundu"
Const strTxtResults = "sonuçlar"
Const strTxtNoSearchResults = "Aramanýzda hiçbir sonuç bulunamadý"
Const strTxtClickHereToRefineSearch = "Tekrar arama yapmak için týklayýnýz."
Const strTxtSearchFor = "Search For"
Const strTxtSearchIn = "Ýçinde Ara"
Const strTxtSearchOn = "Search On"
Const strTxtAllWords = "Tüm Kelimeler"
Const strTxtAnyWords = "Herhangi Kelime"
Const strTxtPhrase = "Phrase"
Const strTxtTopicSubject = "Konu Baþlýðý"
Const strTxtMessageBody = "Mesaj Ýçeriði"
Const strTxtAuthor = "Yazar"
Const strTxtSearchForum = "Forumda Ara"
Const strTxtSortResultsBy = "Sonuçlarý Sýrala"
Const strTxtLastPostTime = "Mesajlarýn Tarihlerine Göre"
Const strTxtTopicStartDate = "Konu Açýlma Tarihine Göre"
Const strTxtSubjectAlphabetically = "Alfabetik Baþlýklara Göre"
Const strTxtNumberViews = "Okunma Sayýsýna Göre"
Const strTxtStartSearch = "Aramaya Baþla"


'printer_friendly_posts.asp
'---------------------------------------------------------------------------------
Const strTxtPrintPage = "Sayfayý Yazdýr"
Const strTxtPrintedFrom = "Printed From"
Const strTxtForumName = "Forum Adý"
Const strTxtForumDiscription = "Forum Açýklamasý"
Const strTxtURL = "URL"
Const strTxtPrintedDate = "Yazdýrýlma Tarihi"
Const strTxtTopic = "Baþlýk"
Const strTxtPostedBy = "Mesajý gönderen"
Const strTxtDatePosted = "Mesaj Tarihi"


'emoticons.asp
'---------------------------------------------------------------------------------
Const strTxtEmoticonSmilies = "Ýfadeler"


'login.asp
'---------------------------------------------------------------------------------
Const strTxtSorryUsernamePasswordIncorrect = "Kullanýcý Adýnýz veya Þifreniz hatalýdýr."
Const strTxtPleaseTryAgain = "Lütfen tekrar deneyin."
Const strTxtUsername = "Kullanýcý Adý"
Const strTxtPassword = "Þifre"
Const strTxtLoginUser = "Giriþ Yap"
Const strTxtClickHereForgottenPass = "Þifrenizi mi unuttunuz?"
Const strTxtErrorUsername = "Kullanýcý Adý \t- Kullanýcý adýnýzý yazýnýz"
Const strTxtErrorPassword = "Þifre \t- Þifrenizi yazýnýz."


'forgotten_password.asp
'---------------------------------------------------------------------------------
Const strTxtForgottenPassword = "Þifremi Unuttum"
Const strTxtNoRecordOfUsername = "Girmiþ olduðunuz kriterlerde hiç sonuç bulunamadý."
Const strTxtNoEmailAddressInProfile = "Üyelik detaylarýnýz E-Posta yoluyla yollanamadý çünkü profilinize E-Posta adresi girilmemiþ."
Const strTxtReregisterForForum = "Foruma tekrar kayýt olmalýsýnýz."
Const strTxtPasswordEmailToYou = "Üyelik detaylarýnýz ve yeni þifreniz E-Posta adresinize gönderildi."
Const strTxtPleaseEnterYourUsername = "Lütfen aþaðýdaki forma Kullanýcý Adýnýzý veya E-Posta adresinizi yazýnýz. Üyelik detaylarýnýz E-Posta adresinize gönderilecektir."
Const strTxtEmailPassword = "Yeni þifre oluþtur"

Const strTxtEmailPasswordRequest = "Üyelik bilgilerinizi kurtarmak için yapýlan istek doðrultusunda bu mail gönderilmiþtir,"
Const strTxtEmailPasswordRequest2 = "Üyelik ve giriþ bilgilerinizi aþaðýda bulabilirsiniz."
Const strTxtEmailPasswordRequest3 = "Foruma dönmek için aþaðýdaki linke týklayýnýz: -"


'forum_password_form.asp
'---------------------------------------------------------------------------------
Const strTxtForumLogin = "Forum Giriþi"
Const strTxtErrorEnterPassword = "Þifre \t- Foruma giriþ için þifrenizi giriniz"
Const strTxtPasswordRequiredForForum = "Bu forumu kullanabilmek için þifreniz olmalýdýr."
Const strTxtForumPasswordIncorrect = "Þifrenizi hatalý girdiniz.."
Const strTxtAutoLogin = "Bu bilgisayarda beni hatýrla (Çerezlerin açýk olmasý gerekmektedir.)"
Const strTxtLoginToForum = "Foruma giriþ yap"


'profile.asp
'---------------------------------------------------------------------------------
Const strTxtNoUserProfileFound = "Bu kullanýcýnýn profil bilgilerine ulaþýlamadý"
Const strTxtRegisteredToViewProfile = "Baþkalarýnýn profillerini görebilmek için forum üyesi olmalýsýnýz."
Const strTxtMemberNo = "Üye No."
Const strTxtEmailAddress = "E-Posta Adresi"
Const strTxtPrivate = "Özel"


'new_topic_form.asp
'---------------------------------------------------------------------------------
Const strTxtPostNewTopic = "Yeni konu oluþtur"
Const strTxtErrorTopicSubject = "Baþlýk \t\t- Yeni konunuz için bir baþlýk giriniz."
Const strTxtForumMemberSuspended = "Forum üyeliðiniz Dondurulduðu veya Aktif olmadýðý için bu özelliði kullanamazsýnýz!"

'edit_post_form.asp
'---------------------------------------------------------------------------------
Const strTxtNoPermissionToEditPost = "Bu mesajý düzenlemeye yetkiniz yoktur!"
Const strTxtReturnForumTopic = "Konuya geri dön"


'email_topic.asp
'---------------------------------------------------------------------------------
Const strTxtEmailTopicToFriend = "Bu konuyu arkadaþýnýza yollayýn"
Const strTxtFriendSentEmail = "Arkadaþýnýzýn E-Posta adresine yollanmýþtýr"
Const strTxtFriendsName = "Arkadaþýnýzýn Adý"
Const strTxtFriendsEmail = "Arkadaþýnýzýn E-Posta Adresi"
Const strTxtYourName = "Sizin Adýnýz"
Const strTxtYourEmail = "Sizin E-Posta Adresiniz"
Const strTxtSendEmail = "Gönder"
Const strTxtMessage = "Mesaj"

Const strTxtEmailFriendMessage = "Düþündümde aþaðýdaki baþlýk ilgini çekebilir"
Const strTxtFrom = "gönderen:"

Const strTxtErrorFrinedsName = "Arkadaþýnýzýn Adý \t- Arkadaþýnýzýn adýný giriniz"
Const strTxtErrorFriendsEmail = "Arkadaþýnýzýn E-Posta Adresi \t- Arkadaþýnýza ait geçerli bir e-posta adresi giriniz."
Const strTxtErrorYourName = "Sizin adýnýz \t- Adýnýzý giriniz"
Const strTxtErrorYourEmail = "Sizin E-Posta Adresiniz \t- Geçerli E-Posta adresinizi giriniz."
Const strTxtErrorEmailMessage = "Mesaj \t- Yollamak istediðiniz mesajý giriniz."



'members.asp
'---------------------------------------------------------------------------------
Const strTxtForumMembersList = "Forum Üye Listesi"
Const strTxtMemberSearch = "Üye Arama"

Const strTxtRegistered = "Kayýt Tarihi"
Const strTxtSend = "Gönder"
Const strTxtNext = "Sonraki"
Const strTxtPrevious = "Önceki"
Const strTxtPage = "Sayfa"

Const strTxtErrorMemberSerach = "Üye Arama\t- Arama için üye kullanýcý adýný yazýnýz."



'register.asp
'---------------------------------------------------------------------------------
Const strTxtRegisterNewUser = "Foruma Kayýt Olun"

Const strTxtProfileUsernameLong = "Forumu kullandýðýnýzda buradaki isim gözükecektir."
Const strTxtRetypePassword = "Þifrenizi yeniden giriniz."
Const strTxtProfileEmailLong = "Zorunlu deðil fakat yazmamanýz halinde Þifre Hatýrlatma, Cevap Bildirimleri gibi özelliklerden faydalanamazsýnýz."
Const strTxtShowHideEmail = "E-Posta adresimi göster"
Const strTxtShowHideEmailLong = "E-Posta adresinizin baþkalarý tarafýndan görünmesini istemiyorsanýz iþaretlemeyin."
Const strTxtSelectCountry = "Ülkenizi Seçiniz"
Const strTxtProfileAutoLogin = "Foruma geri döndüðümde otomatik giriþ yap"
Const strTxtSignature = "Ýmza"
Const strTxtSignatureLong = "Forum Mesajlarýnýzýn altýnda görünmesi için imza giriniz."

Const strTxtErrorUsernameChar = "Kullanýcý Adý \t- Kullanýcý adýnýz en az 2 karakter olmalýdýr."
Const strTxtErrorPasswordChar = "Þifre \t- Þifreniz en az 4 karakter olmalýdýr"
Const strTxtErrorPasswordNoMatch = "Þifre Hatasý\t- Girdiðiniz þifreler birbirleriyle uyuþmuyor"
Const strTxtErrorValidEmail = "E-Posta\t\t- Geçerli ve doðru bir e-posta adresi giriniz."
Const strTxtErrorValidEmailLong = "Eðer E-Posta adresinizi girmek istemiyorsanýz E-Posta alanýný boþ býrakýnýz"
Const strTxtErrorNoEmailToShow = "E-Posta adresi girmeden baþkalarý tarafýndan görülsün olarak iþaretleyemezsiniz!"
Const strTxtErrorSignatureToLong = "Ýmza \t- Ýmzanýz fazla uzun"
Const strTxtUpdateProfile = "Profilimi Güncelle"


Const strTxtUsrenameGone = "Kullanýcý adýnýz baþka bir kullanýcý tarafýndan alýnmýþ, veya bu kullanýcý adýný kullanma izni yok, veya kullanýcý adýnýz 2 karakterden daha az.\n\nLütfen farklý bir kullanýcý adý deneyiniz."
Const strTxtEmailThankYouForRegistering = "Foruma zaman ayýrýp kayýt olduðunuz için teþekkür ederiz."
Const strTxtEmailYouCanNowUseTheForumAt = "Kayýt bilgilerinizi aþaðýda bulabilirsiniz."
Const strTxtEmailForumAt = "forum at"
Const strTxtEmailToThe = "to "


'register_new_user.inc
'---------------------------------------------------------------------------------
Const strTxtEmailAMeesageHasBeenPosted = "A message has been posted on"
Const strTxtEmailClickOnLinkBelowToView = "Mesajý görüntülemek veya cevap yazmak için aþaðýdaki linke týklayýnýz."
Const strTxtEmailAMeesageHasBeenPostedOnForumNum = "A message has been posted in the forum number"


'registration_rules.asp
'---------------------------------------------------------------------------------
Const strTxtForumRulesAndPolicies = "Forum Kurallarý ve Prensipleri"
Const srtTxtAccept = "Kabul"




'New from version 6
'---------------------------------------------------------------------------------
Const strTxtHi = "Hi"
Const strTxtInterestingForumPostOn = "Interesting Forum post on"
Const strTxtForumLostPasswordRequest = "Forum giriþ bilgileri isteði"
Const strTxtLockForum = "Forumu Kilitle"
Const strTxtLockedTopic = "Konuyu Kilitle"
Const strTxtUnLockTopic = "Konu Kilidini Aç"
Const strTxtTopicLocked = "Konu Kilitlenmiþtir."
Const strTxtUnForumLocked = "Forum Kildini Aç"
Const strTxtThisTopicIsLocked = "Bu konu kilitlenmiþtir."
Const strTxtThatYouAskedKeepAnEyeOn = "that you asked us to keep an eye on."
Const strTxtTheTopicIsNowDeleted = "Bu konu silinmiþtir."
Const strTxtOf = "of"
Const strTxtTheTimeNowIs = "Þu anki saat"
Const strTxtYouLastVisitedOn = "En son giriþ tarihiniz"
Const strTxtSendMsg = "Send PM"
Const strTxtSendPrivateMessage = "Özel Mesaj Gönder"
Const strTxtActiveUsers = "Aktif Kullanýcý" ' Kullanýcýlar yazmak mantýksýz
Const strTxtMembers = "Üye" ' Çoðuk kullanmak mantýksýz
Const strTxtEnterTextYouWouldLikeIn = "Enter the text that you would like in"
Const strTxtEmailAddressAlreadyUsed = "Girmiþ olduðunuz eposta adresi baþka bir üye tarafýndan kullanýlýyor."
Const strTxtIP = "IP"
Const strTxtIPLogged = "IP Logged"
Const strTxtPages = "Sayfalar"
Const strTxtCharacterCount = "Karakter Sayacý"
Const strTxtAdmin = "Yönetici"


Const strTxtType = "Grup"
Const strTxtActive = "Aktif"
Const strTxtGuest = "Misafir"
Const strTxtAccountStatus = "Hesap Durumu"
Const strTxtNotActive = "Aktif Deðil"



Const strTxtEmailRequiredForActvation = "Üyeliðinizin aktif olmasý için bir eposta alacaksýnýz."
Const strTxtToActivateYourMembershipFor = "Üyeliðinizi aktif etmek için"
Const strTxtForumClickOnTheLinkBelow = "aþaðýdaki linke týklayýnýz."
Const strTxtForumAdmin = "Forum Admin"
Const strTxtViewLastPost = "Son mesajý göster"
Const strTxtSelectAvatar = "Avatar Seçiniz"
Const strTxtAvatar = "Avatar"
Const strTxtSelectAvatarDetails = "Mesajlarýnýzda gösterilmek üzere küçük resimdir. Listeden küçük resim seçebileceðiniz gibi kendiniz de yükleyebilirsiniz (64 x 64 piksel boyutlarýnda olmalýdýr)"
Const strTxtForumCodesInSignature = "imzanýzda kullanabilirsiniz"

Const strTxtHighPriorityPost = "Duyurular"
Const strTxtPinnedTopic = "Sabit Konu"

Const strTxtOpenForum = "Open Forum"
Const strTxtReadOnly = "Sadece Okunabilir"
Const strTxtPasswordRequired = "Þifre Gereklidir"
Const strTxtNoAccess = "Eriþim Yok"

Const strTxtFont = "Font"
Const strTxtSize = "Boyut"
Const strTxtForumCodes = "BBcode Kodlarýný"

Const strTxtNormal = "Normal Konu"
Const strTxtTopAllForums = "Duyurular (tüm forumlarda)"
Const strTopThisForum = "Duyurular (bu forumda)"


Const strTxtMarkAllPostsAsRead = "Tüm mesajlarý okundu iþaretle"
Const strTxtDeleteCookiesSetByThisForum = "Bu forum için çerezleri sil"


'forum_codes
'---------------------------------------------------------------------------------
Const strTxtYouCanUseForumCodesToFormatText = "Yazýnýza biçim vermek için Forum Kodlarýný kullanabilirsiniz."
Const strTxtTypedForumCode = "Yazmýþ olduðunuz Forum Kodu"
Const strTxtConvetedCode = "Çevrilmiþ kod"
Const strTxtTextFormating = "Yazý biçimlendirme"
Const strTxtImagesAndLinks = "Resimler ve Linkler"
Const strTxtMyLink = "Benim Linkim"
Const strTxtMyEmail = "Benim EPosta adresim"



'insufficient_permission.asp
'---------------------------------------------------------------------------------
Const strTxtAccessDenied = "Eriþim Engellendi"
Const strTxtInsufficientPermison = "Sadece yeterli yetkiye sahip olan kullanýcýlar bu sayfayý görebilir."


'activate.asp
'---------------------------------------------------------------------------------
Const strTxtYourForumMemIsNowActive = "Forum üyeliðiniz aktifleþtirildi."
Const strTxtErrorWithActvation = "Forum üyeliðiniz aktive edilirken bir sorun oluþtu.<br /><br />Lütfen iletiþime geçin "


'register_mail_confirm.asp
'---------------------------------------------------------------------------------
Const strTxtYouShouldReceiveAnEmail = "<strong>Forum üyeliðiniz aktive edilmeli!</strong> <br /><br />Üyelik epostasý kayýt iþleminden bir süre sonra eposta adresinize yollanacaktýr.<br />EPostanýzdaki linke týklayarak Forum Üyeliðinizi aktif hale getirebilirsiniz."
Const strTxtThankYouForRegistering = "Kayýt olduðunuz için teþekkür ederiz"
Const strTxtIfErrorActvatingMembership = "Eðer üyeliðinizi aktifleþtirmede sorun yaþýyorsanýz"


'active_users.asp
'---------------------------------------------------------------------------------
Const strTxtActiveForumUsers = "Aktif Forum Kullanýcýlarý"
Const strTxtAddMToActiveUsersList = "Aktif kullanýcýlar arasýna ekle"
Const strTxtLoggedIn = "Logged In"
Const strTxtLastActive = "Son Aktif"
Const strTxtBrowser = "Tarayýcý"
Const strTxtOS = "Sistem"
Const strTxtMinutes = "dakika"
Const strTxtAnnoymous = "Misafir"



'not_posted.asp
'---------------------------------------------------------------------------------
Const strTxtMessageNotPosted = "Mesaj yollanamadý"
Const strTxtDoublePostingIsNotPermitted = "Çift mesaj göndermek yasaklanmýþtýr, mesajýnýz daha önce gönderildi."
Const strTxtSpammingIsNotPermitted = "Spam yapmak yasaktýr!"
Const strTxtYouHaveExceededNumOfPostAllowed = "Belirli bir süre içindeki maksimum mesaj gönderme sayýnýzý aþtýnýz.<br /><br />Lütfen daha sonra tekrar deneyiniz."
Const strTxtYourMessageNoValidSubjectHeading = "Mesajýnýz geçerli bir baþlýða ve/veya içeriðe sahip deðil."


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtActiveTopics = "Yeni Mesajlar"
Const strTxtLastVisitOn = "Son ziyaret"
Const strTxtLastFifteenMinutes = "Son 15 dakika"
Const strTxtLastThirtyMinutes = "Son 30 dakika"
Const strTxtLastFortyFiveMinutes = "Son 45 dakika"
Const strTxtLastHour = "Son 1 saat"
Const strTxtLastTwoHours = "Son 2 saat"
Const strTxtYesterday = "Dün"
Const strTxtNoActiveTopicsSince = "Belirttiðiniz süre içinde hiç konu bulunamadý."
Const strTxtToDisplay = "gösteriliyor."
Const strTxtThereAreCurrently = "There are currently"



'pm_check.inc
'---------------------------------------------------------------------------------
Const strTxtNewPMsClickToGoNowToPM = "yeni özel mesajýnýz var.\n\nÖzel Mesaj bölümüne ulaþmak için Tamam'ý týklayýn."


'display_forum_topics.inc
'---------------------------------------------------------------------------------
Const strTxtFewYears = "birkaç yýl"
Const strTxtWeek = "hafta"
Const strTxtTwoWeeks = "iki hafta"
Const strTxtMonth = "ay"
Const strTxtTwoMonths = "iki ay"
Const strTxtSixMonths = "6 ay"
Const strTxtYear = "yýl"



Const strTxtHasBeenSentTo = "has been sent to"
Const strTxtCharactersInYourSignatureToLong = "imzanýzdaki karakterler 200'den az olmalýdýr."
Const strTxtSorryYourSearchFoundNoMembers = "Yapmýþ olduðunuz arama kriterlerinde hiç üye bulunamamýþtýr, arama kriterlerinizi gözden geçirdikten sonra tekrar deneyiniz"
Const strTxtCahngeOfEmailReactivateAccount = "Eðer eposta adresinizi deðiþtirirseniz üyeliðinizi tekrar aktif etmek için eposta gönderilecektir."
Const strTxtAddToBuddyList = "Arkadaþ listesine ekle"


'register_mail_confirm.asp
'---------------------------------------------------------------------------------
Const strTxtYourEmailAddressHasBeenChanged = "Eposta adresiniz deðiþtirilmiþtir, <br />forum üyeliðinizi tekrar aktive etmeniz gerekmektedir."
Const strTxtYouShouldReceiveAReactivateEmail = "<strong>Forum üyeliðiniz tekrar aktive edilmelidir!</strong><br /><br />Profilinizdeki adrese bir süre sonra tekrar aktivasyon için eposta gelecektir.<br />Forum üyeliðinizi tekrar aktive etmek için epostanýzdaki linke týklaynýz."


'Preview signature window
'---------------------------------------------------------------------------------
Const strTxtSignaturePreview = "Ýmza Önizleme"
Const strTxtPostedMessage = "Gönderilen Mesaj"



'New from version 7
'---------------------------------------------------------------------------------

Const strTxtMemberlist = "Üye Listesi"
Const strTxtForums = "Forum içinde"
Const strTxtOurUserHavePosted = "Üyelerimizin yollamýþ olduklarý mesajlar: "
Const strTxtInTotalThereAre = "Þu anda forumda "
Const strTxtOnLine = "bulunmaktadýr" 'Aktif yazmak mantýksýz
Const strTxtWeHave = "Toplam"
Const strTxtActivateAccount = "Hesabý aktive et"
Const strTxtSorryYouDoNotHavePermissionToPostInTisForum = "Yeni konu açmak için yetkiniz bulunmamaktadýr."
Const strTxtSorryYouDoNotHavePerimssionToReplyToPostsInThisForum = "Mesajlarý cevaplamak için yetkiniz bulunmamaktadýr."
Const strTxtSorryYouDoNotHavePerimssionToReplyIPBanned = "Mesaj yazamazsýnýz, IP Adresiniz engellenmiþtir.<br />Bunun bir hata olduðunu düþünüyorsanýz Lütfen Forum Adminleriyle iletiþim kurunuz."
Const strTxtLoginSm = "giriþ"
Const strTxtYourProfileHasBeenUpdated = "Profiliniz güncellenmiþtir."
Const strTxtPosted = "Gönderildi:"
Const strTxtBackToTop = "Baþa dön"
Const strTxtNewPassword = "Yeni Þifre"
Const strTxtRetypeNewPassword = "Tekrar Yeni Þifre"
Const strTxtRegards = "Saygýlar"
Const strTxtClickTheLinkBelowToUnsubscribe = "Bu Konuyla ilgili veya bu Forumla iligli artýk eposta almak istemiyorsanýz lütfen aþaðýdaki linke týklayýnýz."
Const strTxtPostsPerDay = "günlük mesaj ortalamasý"
Const strTxtGroup = "Grup"
Const strTxtLastVisit = "Son Ziyaret"
Const strTxtPrivateMessage = "Özel Mesaj"
Const strTxtSorryFunctionNotPermiitedIPBanned = "Bu özellik sizin için kullanýlamaz, IP Adresiniz engellenmiþtir.<br />Bunun bir hata olduðunu düþünüyorsanýz Lütfen Forum Adminlerilye iletiþim kurunuz."
Const strTxtEmailAddressBlocked = "Bu eposta adresi ve alan adý Forum Adminleri tarafýndan engellenmiþtir.<br />Lütfen farklý bir eposta adresi veya eposta alan adý kullanýn."
Const strTxtTopicAdmin = "Konu Ayarlarý"
Const strTxtMovePost = "Mesajý taþý"
Const strTxtPrevTopic = "Önceki Konu"
Const strTxtTheMemberHasBeenDleted = "Üye silinmiþtir."
Const strTxtThisPageWasGeneratedIn = "Bu sayfa"
Const strTxtSeconds = "saniyede yüklenmiþtir."
Const strTxtEditBy = "Düzenleyen"
Const strTxtWrote = "yazdý"
Const strTxtEnable = "Aktif"
Const strTxtToFormatPosts = "mesajý biçimlendirmek için kullanabilirsiniz"
Const strTxtFlashFilesImages = "Adobe Flash"
Const strTxtSessionIDErrorCheckCookiesAreEnabled = "Yetkilendirmeyle ilgili güvenlik hatasý oluþtu.<br /><br />Lütfen tarayýcýnýzýn Çerez özelliðinin açýk olduðundan emin olunuz, sayfanýn bilgisayarýnýzda kayýtlý bir kopyasýný kullanamazsýnýz. Ayrýca IP adresinizi gizleyen Firewall/Proxy gibi programlarýnýzý kontrol ediniz."
Const strTxtName = "Ýsim"
Const strTxtModerators = "Moderatörler"
Const strTxtMore = "devamý..."
Const strTxtNewRegSuspendedCheckBackLater = "Yeni kayýt iþlemi durdurulmuþtur, lütfen daha sonra tekrar kontrol edin."
Const strTxtMoved = "Taþýndý"
Const strTxtNoNameError = "Ýsim \t\t- Lütfen adýnýzý yazýnýz"
Const strTxtHelp = "Yardým"

'PM system
'---------------------------------------------------------------------------------
Const strTxtPrivateMessenger = "Özel Mesajlar"
Const strTxtUnreadMessage = "Okunmamýþ Mesajlar"
Const strTxtReadMessage = "Mesaj Oku"
Const strTxtNew = "yeni"
Const strTxtYouHave = "Gelen Kutunuzda"
Const strTxtNewMsgsInYourInbox = "okunmamýþ mesaj(lar) var!"
Const strTxtNoneSelected = "Seçim yapýlmadý"
Const strTxtAddBuddy = "Arkadaþ Ekle"


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtSelectMember = "Üye Seç"
Const strTxtSelect = "Seç"
Const strTxtNoMatchesFound = "Uyuþan bulunamadý"


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtLastFourHours = "Son 4 saat"
Const strTxtLastSixHours = "Son 6 saat"
Const strTxtLastEightHours = "Son 8 saat"
Const strTxtLastTwelveHours = "Son 12 saat"
Const strTxtLastSixteenHours = "Son 16 saat"


'permissions
'---------------------------------------------------------------------------------
Const strTxtYou = "Sen"
Const strTxtCan = "yetkilisin"
Const strTxtCannot = "yetkili deðilsin"
Const strTxtpostNewTopicsInThisForum = "forumda yeni konu oluþturma"
Const strTxtReplyToTopicsInThisForum = "forumda konulara cevap yazma"
Const strTxtEditYourPostsInThisForum = "forumda mesajlarýný deðiþtirme"
Const strTxtDeleteYourPostsInThisForum = "forumda mesajlarýný silme"
Const strTxtCreatePollsInThisForum = "forumda anket oluþturma"
Const strTxtVoteInPOllsInThisForum = "forumda ankete oy verme"


'register.asp
'---------------------------------------------------------------------------------
Const strTxtRegistrationDetails = "Kayýt Detaylarý"
Const strTxtProfileInformation = "Profil Detaylarý (zorunlu deðil)"
Const strTxtForumPreferences = "Forum Ayarlarý"
Const strTxtICQNumber = "ICQ Numarasý"
Const strTxtAIMAddress = "AIM Adresi"
Const strTxtMSNMessenger = "MSN Messenger Adresi"
Const strTxtYahooMessenger = "Yahoo Messenger Adresi"
Const strTxtOccupation = "Meslek"
Const strTxtInterests = "Ýlgi alanlarý"
Const strTxtDateOfBirth = "Doðum Tarihi"
Const strTxtNotifyMeOfReplies = "Mesajlarýma yazýlan cevaplar için beni bilgilendir"
Const strTxtSendsAnEmailWhenSomeoneRepliesToATopicYouHavePostedIn = "Mesajýna cevap yazýldýðýnda eposta adresinize bilgilendirme postasý gelir. Bu ayarý her mesaj yazýþýnýzda deðiþtirebilirsiniz."
Const strTxtNotifyMeOfPrivateMessages = "Özel mesaj aldýðýmda eposta yoluyla beni bilgilendir"
Const strTxtAlwaysAttachMySignature = "Mesajlarýmda herzaman imza kullan"
Const strTxtEnableTheWindowsIEWYSIWYGPostEditor = "WYSIWYG editörünü etkinleþtir <br /><span class=""smText"">Sadece yeni nesil tarayýcýlarda bu özellik kullanýlabilinir, tarayýcýnýz tarafýndan editör otomatik olarak algýlanýr.</span>"
Const strTxtTimezone = "Forum saatine göre saat dilimi"
Const strTxtPresentServerTimeIs = "Þu anda sunucudaki tarih ve saat: "
Const strTxtDateFormat = "Tarih Formatý"
Const strTxtDayMonthYear = "Gün/Ay/Yýl"
Const strTxtMonthDayYear = "Ay/Gün/Yýl"
Const strTxtYearMonthDay = "Yýl/Ay/Gün"
Const strTxtYearDayMonth = "Yýl/Gün/Ay"
Const strTxtHours = "saatler"
Const strTxtDay = "Gün"
Const strTxtCMonth = "Ay"
Const strTxtCYear = "Yýl"
Const strTxtRealName = "Gerçek isim"
Const strTxtMemberTitle = "Üye baþlýðý"


'Polls
'---------------------------------------------------------------------------------
Const strTxtCreateNewPoll = "Yeni anket oluþtur"
Const strTxtPollQuestion = "Anket&nbsp;Sorusu"
Const strTxtPollChoice = "Anket Þýk"
Const strTxtErrorPollQuestion = "Anket Sorusu \t- Anket için soru belirtiniz"
Const strTxtErrorPollChoice = "Anket þýkký \t- Anket için en az iki tane þýk belirleyiniz"
Const strTxtSorryYouDoNotHavePermissionToCreatePollsForum = "Forumda anket oluþturma yetkiniz bulunmamaktadýr."
Const strTxtAllowMultipleVotes = "Bu ankette birden fazla oy vermeyi etkinleþtir."
Const strTxtMakePollOnlyNoReplies = "Sadece anket oluþtur (cevaplara izin verilmez)"
Const strTxtYourNoValidPoll = "Anketiniz geçerli bir soruyu veya þýklarý içermemektedir."
Const strTxtPoll = "Anket:"
Const strTxtVote = "Oy"
Const strTxtVotes = "Oylar"
Const strTxtCastMyVote = " Oy ver"
Const strTxtPollStatistics = "Anket istatistikleri"
Const strTxtThisTopicIsClosedNoNewVotesAccepted = "Bu anket kapatýlmýþtýr, yeni oylar kabul edilmemektedir"
Const strTxtYouHaveAlreadyVotedInThisPoll = "Daha önce bu ankete oy verdiniz"
Const strTxtThankYouForCastingYourVote = "Oy verdiðiniz için teþekkür ederiz."
Const strsTxYouCanNotNotVoteInThisPoll = "Bu ankete oy veremezsiniz"
Const strTxtYouDidNotSelectAChoiceForYourVote = "Oyunuz sayýlmamýþtýr.\n\nOy vermeniz için herhangi bir þýkký iþaretlemiþ olmanýz lazým."
Const strTxtThisIsAPollOnlyYouCanNotReply = "Sadece anket içindir, mesaj yollayamazsýnýz."


'Email Notify
'---------------------------------------------------------------------------------
Const strTxtWatchThisTopic = "Bu konuyu takip et"
Const strTxtUn = "Un-"
Const strTxtWatchThisForum = "Bu forumu takip et"
Const strTxtYouAreNowBeNotifiedOfPostsInThisForum = "Bu forumdaki tüm mesajlar için eposta yoluyla bilgilendirme alacaksýnýz.\n\nBilgilendirme postalarýný istemiyorsanýz \'Forumu takip etme\' \n butonuna týklayýnýz veya Forum Seçenekleri sayfasýndaki \'EPosta Bilgilendirme\' sayfasýný ziyaret ediniz."
Const strTxtYouAreNowNOTBeNotifiedOfPostsInThisForum = "Bu forumdaki tüm mesajlar için artýk eposta almayacaksýnýz.\n\nBilgilendirme postalarýný istiyorsanýz \'Forumu Takip Et\' \n butonuna týklayýnýz veya Forum Seçenekleri sayfasýndaki \'EPosta Bilgilendirme\' sayfasýný ziyaret ediniz."
Const strTxtYouWillNowBeNotifiedOfAllReplies = "Bu konudaki mesajýnýza gönderilen tüm cevaplar için bilgilendirme postalarý alacaksýnýz.\n\nBilgilendirme postalarýný istemiyorsanýz \'Konuyu takip etme\' \n butonuna týklayýnýz veya Forum Seçenekleri sayfasýndaki \'EPosta Bilgilendirme\' sayfasýný ziyaret ediniz."
Const strTxtYouWillNowNOTBeNotifiedOfAllReplies = "Bu konudaki mesajýnýza gönderilen tüm cevaplar için artýk eposta almayacaksýnýz.\n\nBilgilendirme postalarýný istiyorsanýz \'Konuyu Takip Et\' butonuna týklayýnýz."


'email_messenger.asp
'---------------------------------------------------------------------------------
Const strTxtEmailMessenger = "Email Messenger"
Const strTxtRecipient = "Alýcý"
Const strTxtNoHTMLorForumCodeInEmailBody = "Göndereceðiniz eposta sadece metin tabanlýdýr (HTML kodlarý veya Forum kodlarý kullanýlamaz).<br /><br />Cevaplama adresi eposta adresiniz olacaktýr."
Const strTxtYourEmailHasBeenSentTo = "EPostanýz gönderildi"
Const strTxtYouCanNotEmail = "Eposta gönderemezsiniz"
Const strTxtYouDontHaveAValidEmailAddr = "Profilinizde geçerli bir eposta adresi bulunmamaktadýr."
Const strTxtTheyHaveChoosenToHideThierEmailAddr = "seçilen üyeler eposta adreslerini gizlemiþler."
Const strTxtTheyDontHaveAValidEmailAddr = "seçilen üyelerin profillerinde geçerli bir eposta adresi bulunmamaktadýr."
Const strTxtSendACopyOfThisEmailToMyself = "Gönderilen epostanýn bir kopyasýný kendime gönder"
Const strTxtTheFollowingEmailHasBeenSentToYouBy = "Aþaðýdaki epostayý size gönderen"
Const strTxtFromYourAccountOnThe = "from the forum your participate in on "
Const strTxtIfThisMessageIsAbusive = "Eðer gelen eposta spam ise veya rahatsýzlýk verici ise lütfen webmaster ile veya forum yetkilileri ile iletiþim kurunuz"
Const strTxtIncludeThisEmailAndTheFollowing = "Bu epostayý ve devamýný ekle"
Const strTxtReplyToEmailSetTo = "Lütfen bu epostanýn yanýtlama/cevap adresini belirtiniz"
Const strTxtMessageSent = "Posta gönderildi"



'forum_closed.asp
'---------------------------------------------------------------------------------
Const strTxtForumClosed = "Forum Kapalý"
Const strTxtSorryTheForumsAreClosedForMaintenance = "Bakým çalýþmalarý sebebiyle forum kapalýdýr.<br />Lütfen daha sonra deneyiniz."


'report_post.asp
'---------------------------------------------------------------------------------
Const strTxtReportPost = "Mesajý bildir"
Const strTxtSendReport = "Raporu gönder"
Const strTxtProblemWithPost = "Mesajdaki problem"
Const strTxtPleaseStateProblemWithPost = "Lütfen bu mesajla ilgili sýkýntýnýzý yazýnýz, mesajýn bir kopyasý forum adminlerine ve modaratörlerine gönderilecektir."
Const strTxtTheFollowingReportSubmittedBy = "Aþaðýdaki raporu gönderen"
Const strTxtWhoHasTheFollowingIssue = "who has the following issue with this post"
Const strTxtToViewThePostClickTheLink = "Mesajý görüntülemek için aþaðýdaki linke týklayýnýz"
Const strTxtIssueWithPostOn = "Issue With Post on"
Const strTxtYourReportEmailHasBeenSent = "Epostanýz Forum Adminlerine ve Modaratörlerine gönderilmiþtir."


'New from version 7.5
'---------------------------------------------------------------------------------
Const strTxtQuickLogin = "Hýzlý Giriþ"
Const strTxtThisTopicWasStarted = "Konu açýlma tarihi: "
Const strTxtResendActivationEmail = "Aktivasyon epostasýný tekrar gönder"
Const strTxtNoOfStars = "Yýldýz sayýsý"
Const strTxtOnLine2 = "Aktif"
Const strTxtCode = "Kod"
Const strTxtCodeandFixedWidthData = "Kod ve sabit geniþlik datasý"
Const strTxtQuoting = "Alýntý"
Const strTxtMyCodeData = "Kodum ve sabit geniþlik datasý"
Const strTxtQuotedMessage = "Alýntý yapýlmýþ mesaj"
Const strTxtWithUsername = "Kullanýcý adýyla birlikte"
Const strTxtGo = "Git"
Const strTxtDataBasedOnActiveUsersInTheLastXMinutes = "Bu bilgiler son 20 dakika içinde aktif olan üyeleri kapsar"
Const strTxtSoftwareVersion = "Yazýlým Versiyonu"
Const strTxtForumMembershipNotAct = "Forum üyeliðiniz henüz aktive edilmemiþ!"
Const strTxtMustBeRegisteredToPost = "Mesajlarda sýralama yapabilmeniz için forum üyesi olmanýz gerekmektedir."
Const strTxtMemberCPMenu = "Üye Kontrol Paneli"
Const strTxtYouCanAccessCP = "Forum araçlarý ve Forum Seçeneklerini deðiþtirebilirsiniz "
Const strTxtEditMembersSettings = "Bu üyenin forum seçeneklerini deðiþtir"
Const strTxtSecurityCodeConfirmation = "Güvenlik Kodu Onayý (gerekli)"
Const strTxtUniqueSecurityCode = "Güvenlik Kodu"
Const strTxtEnterCAPTCHAcode = "Lütfen resimde gördüðünüz kodu Güvenlik Kodu alanýna giriniz.<br />Tarayýcýnýzýn çerez desteðinin açýk olmasý gerekmektedir."
Const strTxtErrorSecurityCode = "Güvenlik Kodu \t- Resimde gördüðünüz kodu girmelisiniz"
Const strTxtSecurityCodeDidNotMatch = "Girmiþ olduðunuz güvenlik kodu ile resimdeki kod uyuþmamaktadýr.\n\nYeni bir güvenlik kodu resmi oluþturulmuþtur."

'login_user_test.asp
'---------------------------------------------------------------------------------
Const strTxtSuccessfulLogin = "Giriþ baþarýlý"
Const strTxtSuccessfulLoginReturnToForum = "Baþarýyla giriþ yaptýnýz, lütfen bekleyiniz foruma yönlendiriliyorsunuz"
Const strTxtUnSuccessfulLoginText = "Çerez sorunundan dolayý giriþiniz baþarýsýz olmuþtur. <br /><br />Lütfen tarayýcýnýzýn çerez desteðinin açýk olduðundan ve IP Adresinizin gizli olmadýðýndan emin olunuz."
Const strTxtUnSuccessfulLoginReTry = "Buraya týklayarak foruma giriþi tekrar deneyebilirsiniz."
Const strTxtToActivateYourForumMem = "Forum üyeliðinizin aktif olmasý için kayýt olduktan sonra eposta adresinize gelen linke týklamanýz gerekmektedir."

'email_notify_subscriptions.asp
'---------------------------------------------------------------------------------
Const strTxtEmailNotificationSubscriptions = "EPosta Bilgilendirme"
Const strTxtSelectForumErrorMsg = "Forum Seçiniz\t- Bilgilendirme postalarýný istediðiniz forumu seçiniz"
Const strTxtYouHaveNoSubToEmailNotify = "EPosta ile bilgilendirme talimatýnýz bulunmamaktadýr"
Const strTxtThatYouHaveSubscribedTo = "EPosta bilgilendirme talimatlarýnýzdýr aþaðýdadýr"
Const strTxtUnsusbribe = "Takip etme"
Const strTxtAreYouWantToUnsubscribe = "Bunlarýn takip edilmemesini istediðinizden emin misiniz?"



'New from version 7.51
'---------------------------------------------------------------------------------
Const strTxtSubscribeToForum = "Yeni mesajlarý takip et. (EPosta Bilgilendirme ile)"
Const strTxtSelectForumToSubscribeTo = "Takip etmek istediðiniz forumu seçiniz"


'New from version 7.7
'---------------------------------------------------------------------------------
Const strTxtOnlineStatus = "Online"
Const strTxtOffLine = "Offline"


'New from version 7.8
'---------------------------------------------------------------------------------
Const strTxtConfirmOldPass = "Eski Þifreyi Onayla"
Const strTxtConformOldPassNotMatching = "Þifre doðrulamasý kayýtlarýmýzdaki tanýmlamanýz ile uyuþmuyor.\n\nEðer þifrenizi deðiþtirmek istiyorsanýz lütfen eski þifrenizi doðru giriniz"



'New from version 8.0
'---------------------------------------------------------------------------------
Const strTxtSub = "Alt"
Const strTxtHidden = "Gizli"
Const strTxtHidePost = "Mesajý Gizle"
Const strTxtAreYouSureYouWantToHidePost = "Bu mesajý gizlemek istediðinizden emin misiniz?"
Const strTxtModeratedPost = "Pre-Approved Post"
Const strTxtYouArePostingModeratedForum = "You are posting in a moderated forum."
Const strTxtBeforePostDisplayedAuthorised = "Mesajýnýzýn forumda yayýnlanabilmesi için forum adminleri ve moderatörler tarafýndan onaylanmasý gerekmektedir."
Const strTxtHiddenTopics = "Moderated Topics"
Const strTxtVerifiedBy = "Onaylayan"
Const strTxtYourEmailHasChanged = "EPosta adresiniz"
Const strTxtPleaseUseLinkToReactivate = "olarak deðiþtirildi, lütfen üyeliðinizin tekrar aktivasyonu için linke týklayýnýz"
Const strTxtToday = "Bugün"
Const strTxtPreviewPost = "Mesaj Önizleme"
Const strTxtEmailNotify = "Cevap geldiðinde EPosta ile bilgilendir"
Const strTxtAvatarUpload = "Avatar yükle"
Const strTxtClickOnEmoticonToAdd = "Mesajýnýza eklemek istediðiniz emoticon'a týklayýnýz."
Const strTxtUpdatePost = "Mesajý Güncelle"
Const strTxtShowSignature = "Ýmzamý Göster"
Const strTxtQuickReply = "Hýzlý Cevap"
Const strTxtCategory = "Kategori"
Const strTxtReverseSortOrder = "Tersinden Sýrala"
Const strTxtSendPM = "Özel Mesaj Gönder"
Const strTxtSearchKeywords = "Anahtar sözcükleri ara"
Const strTxtSearchbyKeyWord = "Anahtar sözcüklere göre ara"
Const strTxtSearchbyUserName = "Kullanýcý adýna göre ara (Ýsteðe Baðlý)"
Const strTxtMatch = "Eþleþen"
Const strTxtSearchOptions = "Arama Ayarlarý"
Const strTxtCtrlApple = "('control' veya 'apple' tuþuna basarak birden fazla seçebilirisniz)"
Const strTxtFindPosts = "Mesajlarda ara"
Const strTxtAndNewer = "ve Yeniler"
Const strTxtAndOlder = "ve Eskiler"
Const strTxtAnyDate = "Herhangi bir zaman"
Const strTxtNumberReplies = "Cevaplanma Sayýsýna Göre"
Const strTxtExactMatch = "Tam Eþleþen"
Const strTxtSearhExpiredOrNoPermission = "Bu arama geçerli deðil veya arama yapmaya yetkiniz bulunmamaktadýr"
Const strTxtCreateNewSearch = "Yeni arama oluþtur"
Const strTxtNoSearchResultsFound = "Hiç sonuç bulunamadý"
Const strTxtSearchError = "Arama Hatasý"
Const strTxtSearchWordLengthError = "Aramanýzda 3 karakterden az kelime veya kelimeler var"
Const strTxtIPSearchError = "IP Adresinize izin verilen arama limitini aþtýnýz<br /><br />Lütfen yeni arama yapmadan önce 30sn bekleyiniz"
Const strTxtResultsIn = "Sonuçlar"
Const strTxtSecounds = "saniyede oluþturuldu"
Const strTxtFor = "için"
Const strTxtThisSearchWasProcessed = "This search was processed"
Const strTxtError = "Hata"
Const strTxtReply = "Cevap"
Const strTxtClose = "Kapat"
Const strTxtActiveStats = "Active Stats"
Const strTxtInformation = "Bilgilendirme"
Const strTxtCommunicate = "Ýletiþim"
Const strTxtDisplayResultsAs = "Sonuçlarý þu þekilde göster"
Const strTxtViewPost = "Mesajý göster"
Const strTxtPasswordRequiredViewPost = "Mesajý görüntülemek için þifre gerekli"
Const strTxtNewestPostFirst = "Yeni mesajlar baþta"
Const strTxtOldestPostFirst = "Eski mesajlar baþta"
Const strTxtMessageIcon = "Mesaj ikonu"
Const strTxtSkypeName = "Skype Adý"
Const strTxtLastPostDetailNotHiddenDetails = "Please note:- Last Post details don't include details of hidden posts."
Const strTxtOriginallyPostedBy = "Orjinalini yazan:"
Const strTxtViewingTopic = "Görüntülediði Konu:"
Const strTxtViewingIndex = "Giriþi Görüntülüyor"
Const strTxtForumIndex = "Forum Giriþi"
Const strTxtIndex = "Giriþ"
Const strTxtViewing = "Kiþi Görüntülüyor"
Const strTxtSearchingForums = "Forumlarý Arýyor"
Const strTxtSearchingFor = "Bunu Arýyor"
Const strTxtWritingPrivateMessage = "Özel Mesaj Yazýyor"
Const strTxtViewingPrivateMessage = "Özel Mesaj Görüntülüyor"
Const strTxtEditingPost = "Mesaj Düzenliyor"
Const strTxtWritingReply = "Cevap Yazýyor"
Const strTxtWritingNewPost = "Yeni Mesaj Yazýyor"
Const strTxtCreatingNewPost = "Yeni Anket Oluþturuyor"
Const strTxtWhatsGoingOn = "Forumda Neler Oluyor?"
Const strTxtLoadNewCode = "Yeni Kod Yükle"
Const strTxtApprovePost = "Mesajý Onayla"
Const strTxt3LoginAtteptsMade = "Bu kullanýcý için 3 giriþ denemesi yapýlmýþtýr.<br />Lütfen bilgilerinizi girdikten sonra güvenlik kodunuda girin."
Const strTxtSuspendUser = "Üyeliði Askýya Al"
Const strTxtAdminNotes = "Yönetici/Moderatör Notu"
Const strTxtAdminNotesAbout = "Bu bölüme yazacaðýnýz notu sadece yöneticiler ve moderatörler kiþinin profiline baktýklarýnda görebilir. Üye hakkýnda uyarýlar v.b. yazabilirsiniz(max 250 karakter)"
Const strTxtAge = "Yaþ"
Const strTxtUnknown = "Geçersiz"
Const strTxtSuspended = "Askýya Alýndý"
Const strTxtEmailNewUserRegistered = "Aþaðýda yeni kaydolan üyeler listelenmektedir "
Const strTxtToActivateTheNewMembershipFor = "Yeni üyeliði aktifleþtirmek için "
Const strTxtNewMemberActivation = "Yeni Üye Aktivasyonu"
Const strTxtEmailYouCanNowUseOnceYourAccountIsActivatedTheForumAt = "Giriþ bilgileriniz aþaðýdadýr. Üyeliðiniz Forum Yöneticisi tarafýndan onaylandýktan sonra yeni mesaj gönderebilir, mesajlarý cevaplayabilirsiniz"
Const strTxtYouAdminNeedsToActivateYourMembership = "<strong>Üyeliðinizin Forum Yöneticisi tarafýndan onaylanmasý gerekmektedir!</strong>"
Const strTxtEmailYourForumMembershipIsActivatedThe = "Forum üyeliðiniz þu anda aktifleþtirilmiþtir.Yeni mesaj yazabilir, mesajlara cevap verebilirsiniz."
Const strTxtTheAccountIsNowActive = "Hesap aktifleþtirildi!!"
Const strTxtErrorOccuredActivatingTheAccount = "Hesabýn aktifleþtirilmesi sýrasýnda bir sorunla karþýlaþýldý"
Const strTxtMustBeLoggedInAsAdminActivateAccount = "Yeni üyelerin aktivasyonunu yapabilmek için yönetici olarak giriþ yapmýþ olmanýz gerekmektedir. <br /> Yönetici giriþi yaptýktan sonra e-postadaki linki tekrar týklayýn."
Const strTxtTodaysBirthdays = "Bugün Doðum Günü Olan Üyeler"
Const strTxtCalendar = "Takvim"
Const strTxtEventDate = "Olayýn Tarihi"
Const strTxtEvent = "Olay"
Const strTxtCalendarEvent = "Takvim Olayý"
Const strTxtLast = "Son"
Const strTxtRSS = "RSS"
Const strTxtNewPostFeed = "Yeni Mesaj Linki"
Const strTxtLastTwoDays = "Son 2 Gün"
Const strTxtThisRSSFileIntendedToBeSyndicated = "Bu sayfa RSS okuyucular ve web sayfalarýnda eþ zamanlý gösterim için tasarlanmýþtýr."
Const strTxtCurrentFeedContent = "Güncel link içeriði"
Const strTxtSyndicatedForumContent = "Güncel forum içerigi"
Const strTxtSubscribeNow = "RSS Linkini Al!"
Const strTxtSubscribeWithWebBasedRSS = "seçiminizi týklayýn"
Const strTxtWithOtherReaders = "eðer bilgisayarýnýzda RSS okuyucu yüklü ise "
Const strTxtSelectYourReader = "Okuyucunuzu Seçin"
Const strTxtThisIsAnXMLFeedOf = "XML içerik linki"
Const strTxtDirectLinkToThisPost = "Mesajýn Direkt Linki"
Const strTxtWhatIsAnRSSFeed = "RSS Linki Nedir?"


'New from version 8.02
'---------------------------------------------------------------------------------
Const strTxtSecurityCodeDidNotMatch2 = "Girdiðiniz güvenlik kodu ekranda gösterilen ile ayný deðil."


'New from version 8.05
'---------------------------------------------------------------------------------
Const strTxtPleaseDontForgetYourPassword = "Lütfen þifrenizi unutmayýnýz, þifre veritabanýnda kodlanarak saklandýðý için unuttuðunuz þifreyi geri alma olanaðý yoktur. Unutmanýz durumunda Parolamý Unuttum bölümünden Kullanýcý Adýnýzý ve E-posta adresinizi belirterek, yeni bir þifrenin E-posta adresinize gönderilmesini isteyebilirsiniz."
Const strTxtActivationEmail = "Aktivasyon E-postasý" 
Const strTxtTopicReplyNotification = "Konu Cevap Bildirimi"
Const strTxtUserNameOrEmailAddress = "Kullanýcý Adý veya E-posta Adresi"
Const strTxtAnonymousMembers = "Bilinmeyen Üye"
Const strTxtGuests = "Misafir"
Const strTxtNewPosts = "yeni mesajlar"
Const strTxtNoNewPosts = "yeni mesaj yok"
Const strTxtFullReplyEditor = "Tam Editör"


'New from version 9
'---------------------------------------------------------------------------------
Const strTxtForumHome = "Anasayfa"
Const strTxtNewMessages = "Yeni Mesaj"
Const strTxtsoh = "Chat (12 Online)"
Const strTxtFAQ = "Yardým"
Const strTxtsohbet = "Sohbet"
Const strTxtUnAnsweredTopics = "Cevaplanmamýþ Konular"
Const strTxtShowPosts = "Mesajlarý Göster"
Const strTxtModeratorTools = "Moderatör Araçlarý"
Const strTxtResyncTopicPostCount = "Forum Ýstatistiklerini Güncelle"
Const strTxtAdminControlPanel = "Yönetici Kontrol Paneli"
Const strTxtAdvancedSearch = "Geliþmiþ Arama"
Const strTxtLockTopic = "Konuyu Kilitle"
Const strTxtHideTopic = "Konuyu Gizle"
Const strTxtShowTopic = "Konuyu Göster"
Const strTxtTopicOptions = "Konu Seçenekleri"
Const strTxtForumOptions = "Forum Seçenekleri"
Const strTxtFindMembersPosts = "Üyenin Mesajlarýný Bul"
Const strTxtMembersProfile = "Üye Profili"
Const strTxtVisitMembersHomepage = "Üyenin Web Sitesine Git"
Const strTxtFirstPage = "Ýlk Sayfa"
Const strTxtLastPage = "Son Sayfa"
Const strTxtPostOptions = "Mesaj Seçenekleri"
Const strTxtBlockUsersIP = "IP Engelle"
Const strTxtCreateNewTopic = "Yeni Konu Oluþtur"
Const strTxtNewPoll = "Anket"
Const strTxtControlPanel = "Kontrol Paneli"
Const strTxtSubscriptions = "Abonelikler"
Const strTxtMessenger = "Haberci"
Const strTxtBuddyList = "Arkadaþ Listesi"
Const strTxtProfile2 = "Profil"
Const strTxtSubscribe = "Abone Ol"
Const strTxtMultiplePages = "Birçok Sayfa"
Const strTxtCurrentPage = "Geçerli Sayfa"
Const strTxtRefreshPage = "Sayfayý Yenile"
Const strTxtAnnouncements = "Duyurular"
Const strTxtHiddenTopic = "Düzenlenmiþ Konu"
Const strTxtHot = "Sýcak"
Const strTxtLocked = "Kilitli"
Const strTxtNewPost = "Yeni Mesaj"
Const strTxtPoll2 = "Anket"
Const strTxtSticky = "Sabit"
Const strTxtForumPermissions = "Forum Ýzinleri"
Const strTxtForumWithSubForums = "Forum ile Alt Forum"
Const strTxtPostNewTopic2 = "Yeni Konu Aç"
Const strTxtViewDropDown = "Açýlýr Kutu Gör"
Const strTxtFull = "dolu"
Const strNotYetRegistered = "Üye ol"
Const strTxtNewsletterSubscription = "Haber Aboneliði"
Const strTxtSignupToRecieveNewsletters = "Haberleri almak için üye ol " 
Const strTxtNewsBulletins = "Haber Bültenleri"
Const strTxtPublished = "Yayýmlandý"
Const strTxtStartDate = "Baþlangýç Tarihi"
Const strTxtEndDate = "Bitiþ Tarihi"
Const strTxtNotRequiredForSingleDateEvents = "not required for single date events"
Const strTxtIn = ""
Const strTxtGender = "Cinsiyet"
Const strTxtMale = "Bay"
Const strTxtFemale = "Bayan"



Const strTxtFileAlreadyUploaded = "Ayný isimde dosyayý daha önceden yüklemiþsiniz!"
Const strTxtSelectOrRenameFile = "Lütfen baþka bir dosya seçiniz veya dosyanýn adýný deðiþtirip tekrar deneyin."
Const strTxtAllotedFileSpaceExceeded = "Size ayrýlan dosya alanýný aþtýnýz: "
Const strTxtDeleteFileOrImagesUingCP = "Lütfen Üye Kontrol Panelinizdeki Dosya Yönetimini kullanarak kullanmadýðýnýz dosya ve resimleri silin."



'File Manager
Const strTxtFileManager = "Dosya Yönetimi"
Const strTxtFileName = "Dosya Adý"
Const strTxtFileSize = "Dosya Boyutu"
Const strTxtFileType = "Dosya Türü"
Const strTxtFileExplorer = "Dosyalar"
Const strTxtFileProperties = "Dosya Özellikleri"
Const strTxtFilePreview = "Dosya Önizleme"
Const strTxtAllocatedFileSpace = "Dosya Alaný Ýstatistikleri"
Const strTxtYouHaveUsed  = "Kullandýðýnýz Alan:"
Const strTxtFromYour = "&nbsp;&nbsp;&nbsp;Ýzin Verilen Alan:"
Const strTxtOfAllocatedFileSpace = ""
Const strTxtYourFileSpaceIs = "Dosya alanýnýz"
Const strTxtDownloadFile = "Yüklenen Dosya"
Const strTxtNewUpload = "Yeni Yükleme Yap"
Const strTxtDeleteFile = "Dosya Sil"
Const strTxtRenameFile = "Yeniden Adlandýr"
Const strTxtAreYouSureDeleteFile = "Bu dosyayý silmek istediðinize emin misiniz?"
Const strTxtNoFileSelected = "Seçili dosya yok"
Const strTxtTheFileNowDeleted = "Dosya silindi"
Const strTxtYourFileHasBeenSuccessfullyUploaded = "Dosyanýz baþarýlý bir þekilde yüklendi."
Const strTxtSelectUploadType = "Yüklem Türü Seç"
Const strTxtYouTube = "YouTube"
Const strTxtUploadFolderEmpty = "Yükleme Klasörü Boþ"




'New from version 9.04
'---------------------------------------------------------------------------------

Const strTxtAutologinOnlyAppliesToSession = ""
Const strTxtViewUnreadPost = "Okunmamýþ Mesajlarý Gör"



'New from version 9.51
'---------------------------------------------------------------------------------

Const strTxtPendingApproval = "Onay bekliyor"
Const strTxtThatRequiresApproval = "bunlar onay bekliyor."

Const strTxtMovieProperties = "Movie Özellikleri"
Const strTxtMovieType = "Movie Tipi"
Const strTxtYouTubeFileName = "YouTube Dosya adý"
Const strTxtFlashMovieURL = "Flash Movie URL"



'New from version 9.52
'---------------------------------------------------------------------------------
Const strTxtThroughTheirForumProfileAtLinkBelow = "through their forum profile at the link below."
Const strTxtYouCanNotEmailTisTopicToAFriend = "You can not email this topic to a friend"
Const strTxtToReplyPleaseEmailContact = "To reply to this email contact"
Const strTxtInsertMovie = "Flash Movie Ekle"


'New from version 9.54
'---------------------------------------------------------------------------------
Const strTxtTheEmailFailedToSendPleaseContactAdmin = "EPosta gönderimi baþarýsýz oldu. Lütfen hata mesajýyla birlikte forum adminlerine ulaþon."
Const strTxtFindMember = "Üye bul"
Const strTxtSearchForTopicsThisMemberStarted = "Bu üyenin açmýþ olduðu konularý bul"
Const strTxtMemberName2 = "Üye Adý"
Const strTxtSearchTimeoutPleaseNarrowSearchTryAgain = "Aramanýz zaman aþýmýna uðradý. Lütfen arama kriterlerinizi gözden geçirip tekrar deneyin"


'New from version 9.55
'---------------------------------------------------------------------------------
Const strTxtTheFileFailedTHeSecurityuScanAndHasBeenDeleted = "Dosya güvenlik taramasýný geçemedi ve silindi, dosya içinde zararlý kodlar olabilir."


'New from version 9.56
'---------------------------------------------------------------------------------
Const strTxtShareTopic = "Konuyu Paylaþ"
Const strTxtPostThisTopicTo = "Bu kadar konuyu ilan et"

'New from version 9.61
'---------------------------------------------------------------------------------
Const strTxtSponsor = "Sponsorlar"

'New from version 9.64
'---------------------------------------------------------------------------------
Const strTxtIPAddress = "IP Adresi"


'New from version 9.65
'---------------------------------------------------------------------------------
Const strTxtTranslate = "Çeviri"

'New from version 9.66
'---------------------------------------------------------------------------------
Const strTxtConfirmEmail = "E-Posta Onayla"
Const strTxtErrorConfirmEmail = "E-posta Onaylama Ýçin Alanlar Eþleþmiyor"


'New from version 9.67
'---------------------------------------------------------------------------------
Const strTxtThereMayAlsoBeOtherMessagesPostedOn = "There may also be other messages posted on"
Const strTxtWarningYourSessionHasExpiredRefreshPageFormDataWillBeLost = "UYARI\nOturumunuz zaman aþýmýna uðradý! Tarayýcýnýzdaki 'Yenile\' butonuna basarak sayfayý yenileyin.\n** Girmiþ olduðunuz form datalarý kaybolacaktýr! **"


'New from version 9.70
'---------------------------------------------------------------------------------
Const strTxtNoFollowAppliedToAllLinks = "NoFollow forumdaki tüm linkler için aktif hale getirildi (rel=""nofollow"")"

'New from version 9.71
'---------------------------------------------------------------------------------
Const strTxtViewIn = "Görünüm"
Const strTxtMoble = "Mobil"
Const strTxtClassic = "Klasik"

'New from version 10
'---------------------------------------------------------------------------------
Const strTxtStatus = "Durum"
Const strTxtTheEmailAddressEnteredIsInvalid = "E-Mail adresi hatalý"
Const strTxtMostUsersEverOnlineWas = "Bugüne kadar en fazla online olan kiþi sayýsý"
Const strTxtTypeTheNameOfMemberInBoxBelow = "Bulmak istediðiniz üye adýnýn tamamýný veya bir kýsmýný yazýn.."
Const strTxtSelectNameOfMemberFromDropDownBelow = "Below is a list of members who match your search criteria, select the member you are looking for and click the 'Select' button to insert this members name into the form this box opened from."
Const strTxtCharacters = "karakterler"
Const strTxtMxLFailedLoginAttemptsMade = "More than the maximum failed login attempts have been made on this account.<br />Please enter your details again, including security code."
Const strTxtNumberOfPoints = "Noktalarýn Sayýsý"
Const strTxtPoints = "Puanlar"
Const strTxtPasswordNotComplex = "Þifreniz karýþýk karakterlerden oluþmalýdýr.\nEn az 1 büyük harf, 1 küçük harf ve 1 sayý içermelidir."
Const strTxtLadderGroup = "Merdiven Grup"
Const strTxtNone = "Hiçbiri"
Const strTxtRealNameError = "Gerçek adýnýzý giriniz"
Const strTxtLocationError = "Lokasyon giriniz"
Const strTxtNotRequired = "Gerekli deðil"
Const strTxtRating = "Oylama"
Const strTxtTopicRating = "Baþlýk oylama"
Const strTxtAverage = "Ortalama"
Const strTxtRateTopic = "Konu Oraný"
Const strTxtYouHaveAlreadyRatedThisTopic = "Bu konu için zaten oy kullanmýþtýnýz!"
Const strTxtYouCanNotRateATopicYouStarted = "You can not rate a Topic that you started!"
Const strTxtThankYouForRatingThisTopic = "Oy kullandýnýz için teþekkürler."
Const strTxtExcellent = "Mükemmel"
Const strTxtPoor = "Kötü"
Const strTxtGood = "Iyi"
Const strTxtTerrible = "Korkunç"
Const strTxtRateThisTopicAs = "Konu oraný için burayý týklayýn"
Const strTxtSortOrder = "Sýrala"
Const strTxtMost = "En"
Const strTxtHighestRating = "Yüksek oranlý"
Const strTxtBy1 = "tarafýndan"
Const strTxtMembers2 = "Üyeler"
Const strTxtEvents = "Etkinlikler"
Const strTxtYouAreOnlyPermittedToEditPostWithin = "Mesajý düzenleyebilirsiniz"
Const strTxtChat = "Sohbet"
Const strTxtOnlineMembers = "Online Kullanýcýlar"
Const strTxtReason = "Akýl"
Const strTxtYourMessageWasRejectedByTheSpamFilters = "Mesajýnýz spam filtreleri tarafýndan reddedildi"
Const strYouMustEnterYour = "Girmelisiniz"
Const strTxtViewUnreadPost1 = "Okunmamýþ Mesajlar"
Const strTxtSetAsAnswer = "Cevap olarak ayarla"
Const strTxtUnSetAsAnswer = "Cevap olarak ayarlamaktan kaldýr"
Const strTxtExternalLinkTo = "Dýþarýya baðlantý adresi"
Const strTxtYouMustBeARegisteredMemberAndPostAReplyToViewMessage = "Gizli içeriði görmek için kayýt olup ve konuya mesaj yazmanýz gerekmektedir."
Const strTxtHideContent = "Ýçeriði Gizle"
Const strTxtPostContentHiddenUntilReply = "Gizli içeriði görmek için konuya cevap yazmanýk gerekir."


Const strTxtThanks = "Teþekkürler"
Const strTxtThanked = "Teþekkür Edildi"
Const strTxtYouMustHaveAnActiveMemberAccount  = "Aktif Üye Hesabý olmalýdýr"
Const strTxtYouCanNotThankYourself = "Kendinize teþekkür edemezsiniz."
Const strTxtYouHaveAlreadySaidThanksForThisPost = "Zaten mesaj için teþekkür edildi."
Const strTxtHasBeenThankedForTheirPost = "has been thanked for their Post"
Const strTxtYourPmInboxIsFullPleaseDeleteOldPMs = "Özel Mesaj Gelen Kutusu dolu! Herhangi bir eski veya istenmeyen özel mesajlarýnýzý silin."
Const strTxtShare = "Paylaþ"
Const strTxtShareThisPageOnTheseSites ="Sayfayý Paylaþ"

Const strTxtFacebook = "Facebook"
Const strTxtLinkedIn = "LinkedIn"
Const strTxtTwitter = "Twitter"

Const strTxtAnswer = "Cevap"
Const strTxtResolution = "Çözüm"
Const strTxtOfficialResponse = "Resmi Yanýt"
%>