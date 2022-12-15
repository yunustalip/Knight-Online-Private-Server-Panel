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
Const strTxtWelcome = "Ho� Geldiniz"
Const strTxtAllForums = "T�m Forumlar"
Const strTxtTopics = "Konular"
Const strTxtPosts = "Mesajlar"
Const strTxtLastPost = "Son Mesaj"
Const strTxtPostPreview = "Mesaj �n izleme"
Const strTxtAt = "saat"
Const strTxtBy = "yazan:"
Const strTxtOn = ""
Const strTxtProfile = "Kullan�c� profili"
Const strTxtSearch = "Arama"
Const strTxtQuote = "Al�nt�"
Const strTxtVisit = "Ziyaret"
Const strTxtView = "G�sterim"
Const strTxtHome = "Ana"
Const strTxtHomepage = "Web Siteniz"
Const strTxtEdit = "D�zenle"
Const strTxtDelete = "Sil"
Const strTxtEditProfile = "Profili D�zenle"
Const strTxtLogOff = "��k��"
Const strTxtRegister = "Foruma Kay�t Olun"
Const strTxtLogin = "Giri�"
Const strTxtMembersList = "Forum �yelerini G�ster"
Const strTxtForumLocked = "Forum Kilitli"
Const strTxtSearchTheForum = "Forum Aramas�"
Const strTxtPostReply = "Cevap Yaz"
Const strTxtNewTopic = "Yeni Konu"
Const strTxtNoForums = "G�sterilecek hi� forum yok"
Const strTxtReturnToDiscussionForum = "Foruma Geri D�n"
Const strTxtMustBeRegistered = "Bu �zelli�i kullanabilmek i�in foruma kay�t olmal�s�n�z."
Const strTxtClearForm = "Formu Temizle"
Const strTxtYes = "Evet"
Const strTxtNo = "Hay�r"
Const strTxtForumLockedByAdmim = "Giri� engellendi.<br />Bu Forum Y�netici taraf�ndan kilitlenmi�tir."
Const strTxtRequiredFields = "��aretli alanlar zorunludur"

Const strTxtForumJump = "Foruma Atla"
Const strTxtSelectForum = "Forum Se�iniz"

Const strTxtNoMessageError = "Mesaj Kutusu \t\t- Mesaj alan� bo� olamaz."
Const strTxtErrorDisplayLine = "_______________________________________________________________"
Const strTxtErrorDisplayLine1 = "Formdaki hatalardan dolay� i�leminiz tamamlanamad�."
Const strTxtErrorDisplayLine2 = "L�tfen hatalar� giderdikten sonra tekrar deneyiniz."
Const strTxtErrorDisplayLine3 = "A�a��daki alan(lar) d�zeltilmelidir: -"



'default.asp
'---------------------------------------------------------------------------------
Const strTxtCookies = "Forumu kullanabilmek i�in �erezler'in ve JavaScript'in taray�c�n�zdan a��k olmas� gerekmektedir."
Const strTxtForum = "Forum"
Const strTxtLatestForumPosts = "Son Forum Mesajlar�"
Const strTxtForumStatistics = "Forum �statistikleri"
Const strTxtNoForumPostMade = "Forumda hi� mesaj bulunmamaktad�r"
Const strTxtThereAre = "Toplam"
Const strTxtPostsIn = "Mesaj,"
Const strTxtTopicsIn = "Konu," 'Konular�n i�inde yazmak mant�ks�z
Const strTxtLastPostBy = "Last Post by"
Const strTxtForumMembers = "Forum �yemiz vard�r"
Const strTxtTheNewestForumMember = "En yeni forum �yemiz"


'forum_topics.asp
'---------------------------------------------------------------------------------
Const strTxtThreadStarter = "Konuyu A�anlar"
Const strTxtReplies = "Cevaplar"
Const strTxtViews = "Okunma"
Const strTxtDeleteTopicAlert = "Bu konuyu silmek istedi�inizden emin misiniz?"
Const strTxtDeleteTopic = "Konuyu Sil"
Const strTxtNextTopic = "Sonraki Konu"
Const strTxtLastTopic = "Son Konu"
Const strTxtShowTopics = "Konular� g�ster"
Const strTxtNoTopicsToDisplay = "Bu forumun i�inde g�sterilecek hi� mesaj yok"

Const strTxtAll = "Hepsi"
Const strTxtLastWeek = "Son Hafta"
Const strTxtLastTwoWeeks = "Son iki Hafta"
Const strTxtLastMonth = "Son Ay"
Const strTxtLastTwoMonths = "Son �ki Ay"
Const strTxtLastSixMonths = "Son Alt� Ay"
Const strTxtLastYear = "Son Y�l"


'forum_posts.asp
'---------------------------------------------------------------------------------
Const strTxtLocation = "Konum"
Const strTxtJoined = "Kay�t tarihi"
Const strTxtForumAdministrator = "Forum Y�neticisi"
Const strTxtForumModerator = "Forum Operat�r�"
Const strTxtDeletePostAlert = "Bu mesaj� silmek istedi�inizden emin misiniz?"
Const strTxtEditPost = "Mesaj� D�zenle"
Const strTxtDeletePost = "Mesaj� Sil"
Const strTxtSearchForPosts = "Mesajlarda ara"
Const strTxtSubjectFolder = "Ba�l�k"
Const strTxtPrintVersion = "Yazd�r"
Const strTxtEmailTopic = "Email g�nder"
Const strTxtSorryNoReply = "Eri�im engellendi."
Const strTxtThisForumIsLocked = "Bu forum y�netici taraf�ndan kilitlenmi�tir."
Const strTxtPostAReplyRegister = "E�er bu konuya mesaj yollamak istiyorsan�z �ncelikle"
Const strTxtNeedToRegister = "E�er foruma kay�t olmam��san�z �ncelikle kay�t olmal�s�n�z"
Const strTxtSmRegister = "kay�t"
Const strTxtNoThreads = "Bu konuya ait hi�bir mesaj bulunamad�"
Const strTxtNotGiven = "Girilmedi"


'search_form.asp
'---------------------------------------------------------------------------------
Const strTxtSearchFormError = "Arama\t- L�tfen aranacak kelimeyi yaz�n"


'search.asp
'---------------------------------------------------------------------------------
Const strTxtSearchResults = "Arama Sonu�lar�"
Const strTxtHasFound = "bulundu"
Const strTxtResults = "sonu�lar"
Const strTxtNoSearchResults = "Araman�zda hi�bir sonu� bulunamad�"
Const strTxtClickHereToRefineSearch = "Tekrar arama yapmak i�in t�klay�n�z."
Const strTxtSearchFor = "Search For"
Const strTxtSearchIn = "��inde Ara"
Const strTxtSearchOn = "Search On"
Const strTxtAllWords = "T�m Kelimeler"
Const strTxtAnyWords = "Herhangi Kelime"
Const strTxtPhrase = "Phrase"
Const strTxtTopicSubject = "Konu Ba�l���"
Const strTxtMessageBody = "Mesaj ��eri�i"
Const strTxtAuthor = "Yazar"
Const strTxtSearchForum = "Forumda Ara"
Const strTxtSortResultsBy = "Sonu�lar� S�rala"
Const strTxtLastPostTime = "Mesajlar�n Tarihlerine G�re"
Const strTxtTopicStartDate = "Konu A��lma Tarihine G�re"
Const strTxtSubjectAlphabetically = "Alfabetik Ba�l�klara G�re"
Const strTxtNumberViews = "Okunma Say�s�na G�re"
Const strTxtStartSearch = "Aramaya Ba�la"


'printer_friendly_posts.asp
'---------------------------------------------------------------------------------
Const strTxtPrintPage = "Sayfay� Yazd�r"
Const strTxtPrintedFrom = "Printed From"
Const strTxtForumName = "Forum Ad�"
Const strTxtForumDiscription = "Forum A��klamas�"
Const strTxtURL = "URL"
Const strTxtPrintedDate = "Yazd�r�lma Tarihi"
Const strTxtTopic = "Ba�l�k"
Const strTxtPostedBy = "Mesaj� g�nderen"
Const strTxtDatePosted = "Mesaj Tarihi"


'emoticons.asp
'---------------------------------------------------------------------------------
Const strTxtEmoticonSmilies = "�fadeler"


'login.asp
'---------------------------------------------------------------------------------
Const strTxtSorryUsernamePasswordIncorrect = "Kullan�c� Ad�n�z veya �ifreniz hatal�d�r."
Const strTxtPleaseTryAgain = "L�tfen tekrar deneyin."
Const strTxtUsername = "Kullan�c� Ad�"
Const strTxtPassword = "�ifre"
Const strTxtLoginUser = "Giri� Yap"
Const strTxtClickHereForgottenPass = "�ifrenizi mi unuttunuz?"
Const strTxtErrorUsername = "Kullan�c� Ad� \t- Kullan�c� ad�n�z� yaz�n�z"
Const strTxtErrorPassword = "�ifre \t- �ifrenizi yaz�n�z."


'forgotten_password.asp
'---------------------------------------------------------------------------------
Const strTxtForgottenPassword = "�ifremi Unuttum"
Const strTxtNoRecordOfUsername = "Girmi� oldu�unuz kriterlerde hi� sonu� bulunamad�."
Const strTxtNoEmailAddressInProfile = "�yelik detaylar�n�z E-Posta yoluyla yollanamad� ��nk� profilinize E-Posta adresi girilmemi�."
Const strTxtReregisterForForum = "Foruma tekrar kay�t olmal�s�n�z."
Const strTxtPasswordEmailToYou = "�yelik detaylar�n�z ve yeni �ifreniz E-Posta adresinize g�nderildi."
Const strTxtPleaseEnterYourUsername = "L�tfen a�a��daki forma Kullan�c� Ad�n�z� veya E-Posta adresinizi yaz�n�z. �yelik detaylar�n�z E-Posta adresinize g�nderilecektir."
Const strTxtEmailPassword = "Yeni �ifre olu�tur"

Const strTxtEmailPasswordRequest = "�yelik bilgilerinizi kurtarmak i�in yap�lan istek do�rultusunda bu mail g�nderilmi�tir,"
Const strTxtEmailPasswordRequest2 = "�yelik ve giri� bilgilerinizi a�a��da bulabilirsiniz."
Const strTxtEmailPasswordRequest3 = "Foruma d�nmek i�in a�a��daki linke t�klay�n�z: -"


'forum_password_form.asp
'---------------------------------------------------------------------------------
Const strTxtForumLogin = "Forum Giri�i"
Const strTxtErrorEnterPassword = "�ifre \t- Foruma giri� i�in �ifrenizi giriniz"
Const strTxtPasswordRequiredForForum = "Bu forumu kullanabilmek i�in �ifreniz olmal�d�r."
Const strTxtForumPasswordIncorrect = "�ifrenizi hatal� girdiniz.."
Const strTxtAutoLogin = "Bu bilgisayarda beni hat�rla (�erezlerin a��k olmas� gerekmektedir.)"
Const strTxtLoginToForum = "Foruma giri� yap"


'profile.asp
'---------------------------------------------------------------------------------
Const strTxtNoUserProfileFound = "Bu kullan�c�n�n profil bilgilerine ula��lamad�"
Const strTxtRegisteredToViewProfile = "Ba�kalar�n�n profillerini g�rebilmek i�in forum �yesi olmal�s�n�z."
Const strTxtMemberNo = "�ye No."
Const strTxtEmailAddress = "E-Posta Adresi"
Const strTxtPrivate = "�zel"


'new_topic_form.asp
'---------------------------------------------------------------------------------
Const strTxtPostNewTopic = "Yeni konu olu�tur"
Const strTxtErrorTopicSubject = "Ba�l�k \t\t- Yeni konunuz i�in bir ba�l�k giriniz."
Const strTxtForumMemberSuspended = "Forum �yeli�iniz Donduruldu�u veya Aktif olmad��� i�in bu �zelli�i kullanamazs�n�z!"

'edit_post_form.asp
'---------------------------------------------------------------------------------
Const strTxtNoPermissionToEditPost = "Bu mesaj� d�zenlemeye yetkiniz yoktur!"
Const strTxtReturnForumTopic = "Konuya geri d�n"


'email_topic.asp
'---------------------------------------------------------------------------------
Const strTxtEmailTopicToFriend = "Bu konuyu arkada��n�za yollay�n"
Const strTxtFriendSentEmail = "Arkada��n�z�n E-Posta adresine yollanm��t�r"
Const strTxtFriendsName = "Arkada��n�z�n Ad�"
Const strTxtFriendsEmail = "Arkada��n�z�n E-Posta Adresi"
Const strTxtYourName = "Sizin Ad�n�z"
Const strTxtYourEmail = "Sizin E-Posta Adresiniz"
Const strTxtSendEmail = "G�nder"
Const strTxtMessage = "Mesaj"

Const strTxtEmailFriendMessage = "D���nd�mde a�a��daki ba�l�k ilgini �ekebilir"
Const strTxtFrom = "g�nderen:"

Const strTxtErrorFrinedsName = "Arkada��n�z�n Ad� \t- Arkada��n�z�n ad�n� giriniz"
Const strTxtErrorFriendsEmail = "Arkada��n�z�n E-Posta Adresi \t- Arkada��n�za ait ge�erli bir e-posta adresi giriniz."
Const strTxtErrorYourName = "Sizin ad�n�z \t- Ad�n�z� giriniz"
Const strTxtErrorYourEmail = "Sizin E-Posta Adresiniz \t- Ge�erli E-Posta adresinizi giriniz."
Const strTxtErrorEmailMessage = "Mesaj \t- Yollamak istedi�iniz mesaj� giriniz."



'members.asp
'---------------------------------------------------------------------------------
Const strTxtForumMembersList = "Forum �ye Listesi"
Const strTxtMemberSearch = "�ye Arama"

Const strTxtRegistered = "Kay�t Tarihi"
Const strTxtSend = "G�nder"
Const strTxtNext = "Sonraki"
Const strTxtPrevious = "�nceki"
Const strTxtPage = "Sayfa"

Const strTxtErrorMemberSerach = "�ye Arama\t- Arama i�in �ye kullan�c� ad�n� yaz�n�z."



'register.asp
'---------------------------------------------------------------------------------
Const strTxtRegisterNewUser = "Foruma Kay�t Olun"

Const strTxtProfileUsernameLong = "Forumu kulland���n�zda buradaki isim g�z�kecektir."
Const strTxtRetypePassword = "�ifrenizi yeniden giriniz."
Const strTxtProfileEmailLong = "Zorunlu de�il fakat yazmaman�z halinde �ifre Hat�rlatma, Cevap Bildirimleri gibi �zelliklerden faydalanamazs�n�z."
Const strTxtShowHideEmail = "E-Posta adresimi g�ster"
Const strTxtShowHideEmailLong = "E-Posta adresinizin ba�kalar� taraf�ndan g�r�nmesini istemiyorsan�z i�aretlemeyin."
Const strTxtSelectCountry = "�lkenizi Se�iniz"
Const strTxtProfileAutoLogin = "Foruma geri d�nd���mde otomatik giri� yap"
Const strTxtSignature = "�mza"
Const strTxtSignatureLong = "Forum Mesajlar�n�z�n alt�nda g�r�nmesi i�in imza giriniz."

Const strTxtErrorUsernameChar = "Kullan�c� Ad� \t- Kullan�c� ad�n�z en az 2 karakter olmal�d�r."
Const strTxtErrorPasswordChar = "�ifre \t- �ifreniz en az 4 karakter olmal�d�r"
Const strTxtErrorPasswordNoMatch = "�ifre Hatas�\t- Girdi�iniz �ifreler birbirleriyle uyu�muyor"
Const strTxtErrorValidEmail = "E-Posta\t\t- Ge�erli ve do�ru bir e-posta adresi giriniz."
Const strTxtErrorValidEmailLong = "E�er E-Posta adresinizi girmek istemiyorsan�z E-Posta alan�n� bo� b�rak�n�z"
Const strTxtErrorNoEmailToShow = "E-Posta adresi girmeden ba�kalar� taraf�ndan g�r�ls�n olarak i�aretleyemezsiniz!"
Const strTxtErrorSignatureToLong = "�mza \t- �mzan�z fazla uzun"
Const strTxtUpdateProfile = "Profilimi G�ncelle"


Const strTxtUsrenameGone = "Kullan�c� ad�n�z ba�ka bir kullan�c� taraf�ndan al�nm��, veya bu kullan�c� ad�n� kullanma izni yok, veya kullan�c� ad�n�z 2 karakterden daha az.\n\nL�tfen farkl� bir kullan�c� ad� deneyiniz."
Const strTxtEmailThankYouForRegistering = "Foruma zaman ay�r�p kay�t oldu�unuz i�in te�ekk�r ederiz."
Const strTxtEmailYouCanNowUseTheForumAt = "Kay�t bilgilerinizi a�a��da bulabilirsiniz."
Const strTxtEmailForumAt = "forum at"
Const strTxtEmailToThe = "to "


'register_new_user.inc
'---------------------------------------------------------------------------------
Const strTxtEmailAMeesageHasBeenPosted = "A message has been posted on"
Const strTxtEmailClickOnLinkBelowToView = "Mesaj� g�r�nt�lemek veya cevap yazmak i�in a�a��daki linke t�klay�n�z."
Const strTxtEmailAMeesageHasBeenPostedOnForumNum = "A message has been posted in the forum number"


'registration_rules.asp
'---------------------------------------------------------------------------------
Const strTxtForumRulesAndPolicies = "Forum Kurallar� ve Prensipleri"
Const srtTxtAccept = "Kabul"




'New from version 6
'---------------------------------------------------------------------------------
Const strTxtHi = "Hi"
Const strTxtInterestingForumPostOn = "Interesting Forum post on"
Const strTxtForumLostPasswordRequest = "Forum giri� bilgileri iste�i"
Const strTxtLockForum = "Forumu Kilitle"
Const strTxtLockedTopic = "Konuyu Kilitle"
Const strTxtUnLockTopic = "Konu Kilidini A�"
Const strTxtTopicLocked = "Konu Kilitlenmi�tir."
Const strTxtUnForumLocked = "Forum Kildini A�"
Const strTxtThisTopicIsLocked = "Bu konu kilitlenmi�tir."
Const strTxtThatYouAskedKeepAnEyeOn = "that you asked us to keep an eye on."
Const strTxtTheTopicIsNowDeleted = "Bu konu silinmi�tir."
Const strTxtOf = "of"
Const strTxtTheTimeNowIs = "�u anki saat"
Const strTxtYouLastVisitedOn = "En son giri� tarihiniz"
Const strTxtSendMsg = "Send PM"
Const strTxtSendPrivateMessage = "�zel Mesaj G�nder"
Const strTxtActiveUsers = "Aktif Kullan�c�" ' Kullan�c�lar yazmak mant�ks�z
Const strTxtMembers = "�ye" ' �o�uk kullanmak mant�ks�z
Const strTxtEnterTextYouWouldLikeIn = "Enter the text that you would like in"
Const strTxtEmailAddressAlreadyUsed = "Girmi� oldu�unuz eposta adresi ba�ka bir �ye taraf�ndan kullan�l�yor."
Const strTxtIP = "IP"
Const strTxtIPLogged = "IP Logged"
Const strTxtPages = "Sayfalar"
Const strTxtCharacterCount = "Karakter Sayac�"
Const strTxtAdmin = "Y�netici"


Const strTxtType = "Grup"
Const strTxtActive = "Aktif"
Const strTxtGuest = "Misafir"
Const strTxtAccountStatus = "Hesap Durumu"
Const strTxtNotActive = "Aktif De�il"



Const strTxtEmailRequiredForActvation = "�yeli�inizin aktif olmas� i�in bir eposta alacaks�n�z."
Const strTxtToActivateYourMembershipFor = "�yeli�inizi aktif etmek i�in"
Const strTxtForumClickOnTheLinkBelow = "a�a��daki linke t�klay�n�z."
Const strTxtForumAdmin = "Forum Admin"
Const strTxtViewLastPost = "Son mesaj� g�ster"
Const strTxtSelectAvatar = "Avatar Se�iniz"
Const strTxtAvatar = "Avatar"
Const strTxtSelectAvatarDetails = "Mesajlar�n�zda g�sterilmek �zere k���k resimdir. Listeden k���k resim se�ebilece�iniz gibi kendiniz de y�kleyebilirsiniz (64 x 64 piksel boyutlar�nda olmal�d�r)"
Const strTxtForumCodesInSignature = "imzan�zda kullanabilirsiniz"

Const strTxtHighPriorityPost = "Duyurular"
Const strTxtPinnedTopic = "Sabit Konu"

Const strTxtOpenForum = "Open Forum"
Const strTxtReadOnly = "Sadece Okunabilir"
Const strTxtPasswordRequired = "�ifre Gereklidir"
Const strTxtNoAccess = "Eri�im Yok"

Const strTxtFont = "Font"
Const strTxtSize = "Boyut"
Const strTxtForumCodes = "BBcode Kodlar�n�"

Const strTxtNormal = "Normal Konu"
Const strTxtTopAllForums = "Duyurular (t�m forumlarda)"
Const strTopThisForum = "Duyurular (bu forumda)"


Const strTxtMarkAllPostsAsRead = "T�m mesajlar� okundu i�aretle"
Const strTxtDeleteCookiesSetByThisForum = "Bu forum i�in �erezleri sil"


'forum_codes
'---------------------------------------------------------------------------------
Const strTxtYouCanUseForumCodesToFormatText = "Yaz�n�za bi�im vermek i�in Forum Kodlar�n� kullanabilirsiniz."
Const strTxtTypedForumCode = "Yazm�� oldu�unuz Forum Kodu"
Const strTxtConvetedCode = "�evrilmi� kod"
Const strTxtTextFormating = "Yaz� bi�imlendirme"
Const strTxtImagesAndLinks = "Resimler ve Linkler"
Const strTxtMyLink = "Benim Linkim"
Const strTxtMyEmail = "Benim EPosta adresim"



'insufficient_permission.asp
'---------------------------------------------------------------------------------
Const strTxtAccessDenied = "Eri�im Engellendi"
Const strTxtInsufficientPermison = "Sadece yeterli yetkiye sahip olan kullan�c�lar bu sayfay� g�rebilir."


'activate.asp
'---------------------------------------------------------------------------------
Const strTxtYourForumMemIsNowActive = "Forum �yeli�iniz aktifle�tirildi."
Const strTxtErrorWithActvation = "Forum �yeli�iniz aktive edilirken bir sorun olu�tu.<br /><br />L�tfen ileti�ime ge�in "


'register_mail_confirm.asp
'---------------------------------------------------------------------------------
Const strTxtYouShouldReceiveAnEmail = "<strong>Forum �yeli�iniz aktive edilmeli!</strong> <br /><br />�yelik epostas� kay�t i�leminden bir s�re sonra eposta adresinize yollanacakt�r.<br />EPostan�zdaki linke t�klayarak Forum �yeli�inizi aktif hale getirebilirsiniz."
Const strTxtThankYouForRegistering = "Kay�t oldu�unuz i�in te�ekk�r ederiz"
Const strTxtIfErrorActvatingMembership = "E�er �yeli�inizi aktifle�tirmede sorun ya��yorsan�z"


'active_users.asp
'---------------------------------------------------------------------------------
Const strTxtActiveForumUsers = "Aktif Forum Kullan�c�lar�"
Const strTxtAddMToActiveUsersList = "Aktif kullan�c�lar aras�na ekle"
Const strTxtLoggedIn = "Logged In"
Const strTxtLastActive = "Son Aktif"
Const strTxtBrowser = "Taray�c�"
Const strTxtOS = "Sistem"
Const strTxtMinutes = "dakika"
Const strTxtAnnoymous = "Misafir"



'not_posted.asp
'---------------------------------------------------------------------------------
Const strTxtMessageNotPosted = "Mesaj yollanamad�"
Const strTxtDoublePostingIsNotPermitted = "�ift mesaj g�ndermek yasaklanm��t�r, mesaj�n�z daha �nce g�nderildi."
Const strTxtSpammingIsNotPermitted = "Spam yapmak yasakt�r!"
Const strTxtYouHaveExceededNumOfPostAllowed = "Belirli bir s�re i�indeki maksimum mesaj g�nderme say�n�z� a�t�n�z.<br /><br />L�tfen daha sonra tekrar deneyiniz."
Const strTxtYourMessageNoValidSubjectHeading = "Mesaj�n�z ge�erli bir ba�l��a ve/veya i�eri�e sahip de�il."


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtActiveTopics = "Yeni Mesajlar"
Const strTxtLastVisitOn = "Son ziyaret"
Const strTxtLastFifteenMinutes = "Son 15 dakika"
Const strTxtLastThirtyMinutes = "Son 30 dakika"
Const strTxtLastFortyFiveMinutes = "Son 45 dakika"
Const strTxtLastHour = "Son 1 saat"
Const strTxtLastTwoHours = "Son 2 saat"
Const strTxtYesterday = "D�n"
Const strTxtNoActiveTopicsSince = "Belirtti�iniz s�re i�inde hi� konu bulunamad�."
Const strTxtToDisplay = "g�steriliyor."
Const strTxtThereAreCurrently = "There are currently"



'pm_check.inc
'---------------------------------------------------------------------------------
Const strTxtNewPMsClickToGoNowToPM = "yeni �zel mesaj�n�z var.\n\n�zel Mesaj b�l�m�ne ula�mak i�in Tamam'� t�klay�n."


'display_forum_topics.inc
'---------------------------------------------------------------------------------
Const strTxtFewYears = "birka� y�l"
Const strTxtWeek = "hafta"
Const strTxtTwoWeeks = "iki hafta"
Const strTxtMonth = "ay"
Const strTxtTwoMonths = "iki ay"
Const strTxtSixMonths = "6 ay"
Const strTxtYear = "y�l"



Const strTxtHasBeenSentTo = "has been sent to"
Const strTxtCharactersInYourSignatureToLong = "imzan�zdaki karakterler 200'den az olmal�d�r."
Const strTxtSorryYourSearchFoundNoMembers = "Yapm�� oldu�unuz arama kriterlerinde hi� �ye bulunamam��t�r, arama kriterlerinizi g�zden ge�irdikten sonra tekrar deneyiniz"
Const strTxtCahngeOfEmailReactivateAccount = "E�er eposta adresinizi de�i�tirirseniz �yeli�inizi tekrar aktif etmek i�in eposta g�nderilecektir."
Const strTxtAddToBuddyList = "Arkada� listesine ekle"


'register_mail_confirm.asp
'---------------------------------------------------------------------------------
Const strTxtYourEmailAddressHasBeenChanged = "Eposta adresiniz de�i�tirilmi�tir, <br />forum �yeli�inizi tekrar aktive etmeniz gerekmektedir."
Const strTxtYouShouldReceiveAReactivateEmail = "<strong>Forum �yeli�iniz tekrar aktive edilmelidir!</strong><br /><br />Profilinizdeki adrese bir s�re sonra tekrar aktivasyon i�in eposta gelecektir.<br />Forum �yeli�inizi tekrar aktive etmek i�in epostan�zdaki linke t�klayn�z."


'Preview signature window
'---------------------------------------------------------------------------------
Const strTxtSignaturePreview = "�mza �nizleme"
Const strTxtPostedMessage = "G�nderilen Mesaj"



'New from version 7
'---------------------------------------------------------------------------------

Const strTxtMemberlist = "�ye Listesi"
Const strTxtForums = "Forum i�inde"
Const strTxtOurUserHavePosted = "�yelerimizin yollam�� olduklar� mesajlar: "
Const strTxtInTotalThereAre = "�u anda forumda "
Const strTxtOnLine = "bulunmaktad�r" 'Aktif yazmak mant�ks�z
Const strTxtWeHave = "Toplam"
Const strTxtActivateAccount = "Hesab� aktive et"
Const strTxtSorryYouDoNotHavePermissionToPostInTisForum = "Yeni konu a�mak i�in yetkiniz bulunmamaktad�r."
Const strTxtSorryYouDoNotHavePerimssionToReplyToPostsInThisForum = "Mesajlar� cevaplamak i�in yetkiniz bulunmamaktad�r."
Const strTxtSorryYouDoNotHavePerimssionToReplyIPBanned = "Mesaj yazamazs�n�z, IP Adresiniz engellenmi�tir.<br />Bunun bir hata oldu�unu d���n�yorsan�z L�tfen Forum Adminleriyle ileti�im kurunuz."
Const strTxtLoginSm = "giri�"
Const strTxtYourProfileHasBeenUpdated = "Profiliniz g�ncellenmi�tir."
Const strTxtPosted = "G�nderildi:"
Const strTxtBackToTop = "Ba�a d�n"
Const strTxtNewPassword = "Yeni �ifre"
Const strTxtRetypeNewPassword = "Tekrar Yeni �ifre"
Const strTxtRegards = "Sayg�lar"
Const strTxtClickTheLinkBelowToUnsubscribe = "Bu Konuyla ilgili veya bu Forumla iligli art�k eposta almak istemiyorsan�z l�tfen a�a��daki linke t�klay�n�z."
Const strTxtPostsPerDay = "g�nl�k mesaj ortalamas�"
Const strTxtGroup = "Grup"
Const strTxtLastVisit = "Son Ziyaret"
Const strTxtPrivateMessage = "�zel Mesaj"
Const strTxtSorryFunctionNotPermiitedIPBanned = "Bu �zellik sizin i�in kullan�lamaz, IP Adresiniz engellenmi�tir.<br />Bunun bir hata oldu�unu d���n�yorsan�z L�tfen Forum Adminlerilye ileti�im kurunuz."
Const strTxtEmailAddressBlocked = "Bu eposta adresi ve alan ad� Forum Adminleri taraf�ndan engellenmi�tir.<br />L�tfen farkl� bir eposta adresi veya eposta alan ad� kullan�n."
Const strTxtTopicAdmin = "Konu Ayarlar�"
Const strTxtMovePost = "Mesaj� ta��"
Const strTxtPrevTopic = "�nceki Konu"
Const strTxtTheMemberHasBeenDleted = "�ye silinmi�tir."
Const strTxtThisPageWasGeneratedIn = "Bu sayfa"
Const strTxtSeconds = "saniyede y�klenmi�tir."
Const strTxtEditBy = "D�zenleyen"
Const strTxtWrote = "yazd�"
Const strTxtEnable = "Aktif"
Const strTxtToFormatPosts = "mesaj� bi�imlendirmek i�in kullanabilirsiniz"
Const strTxtFlashFilesImages = "Adobe Flash"
Const strTxtSessionIDErrorCheckCookiesAreEnabled = "Yetkilendirmeyle ilgili g�venlik hatas� olu�tu.<br /><br />L�tfen taray�c�n�z�n �erez �zelli�inin a��k oldu�undan emin olunuz, sayfan�n bilgisayar�n�zda kay�tl� bir kopyas�n� kullanamazs�n�z. Ayr�ca IP adresinizi gizleyen Firewall/Proxy gibi programlar�n�z� kontrol ediniz."
Const strTxtName = "�sim"
Const strTxtModerators = "Moderat�rler"
Const strTxtMore = "devam�..."
Const strTxtNewRegSuspendedCheckBackLater = "Yeni kay�t i�lemi durdurulmu�tur, l�tfen daha sonra tekrar kontrol edin."
Const strTxtMoved = "Ta��nd�"
Const strTxtNoNameError = "�sim \t\t- L�tfen ad�n�z� yaz�n�z"
Const strTxtHelp = "Yard�m"

'PM system
'---------------------------------------------------------------------------------
Const strTxtPrivateMessenger = "�zel Mesajlar"
Const strTxtUnreadMessage = "Okunmam�� Mesajlar"
Const strTxtReadMessage = "Mesaj Oku"
Const strTxtNew = "yeni"
Const strTxtYouHave = "Gelen Kutunuzda"
Const strTxtNewMsgsInYourInbox = "okunmam�� mesaj(lar) var!"
Const strTxtNoneSelected = "Se�im yap�lmad�"
Const strTxtAddBuddy = "Arkada� Ekle"


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtSelectMember = "�ye Se�"
Const strTxtSelect = "Se�"
Const strTxtNoMatchesFound = "Uyu�an bulunamad�"


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
Const strTxtCannot = "yetkili de�ilsin"
Const strTxtpostNewTopicsInThisForum = "forumda yeni konu olu�turma"
Const strTxtReplyToTopicsInThisForum = "forumda konulara cevap yazma"
Const strTxtEditYourPostsInThisForum = "forumda mesajlar�n� de�i�tirme"
Const strTxtDeleteYourPostsInThisForum = "forumda mesajlar�n� silme"
Const strTxtCreatePollsInThisForum = "forumda anket olu�turma"
Const strTxtVoteInPOllsInThisForum = "forumda ankete oy verme"


'register.asp
'---------------------------------------------------------------------------------
Const strTxtRegistrationDetails = "Kay�t Detaylar�"
Const strTxtProfileInformation = "Profil Detaylar� (zorunlu de�il)"
Const strTxtForumPreferences = "Forum Ayarlar�"
Const strTxtICQNumber = "ICQ Numaras�"
Const strTxtAIMAddress = "AIM Adresi"
Const strTxtMSNMessenger = "MSN Messenger Adresi"
Const strTxtYahooMessenger = "Yahoo Messenger Adresi"
Const strTxtOccupation = "Meslek"
Const strTxtInterests = "�lgi alanlar�"
Const strTxtDateOfBirth = "Do�um Tarihi"
Const strTxtNotifyMeOfReplies = "Mesajlar�ma yaz�lan cevaplar i�in beni bilgilendir"
Const strTxtSendsAnEmailWhenSomeoneRepliesToATopicYouHavePostedIn = "Mesaj�na cevap yaz�ld���nda eposta adresinize bilgilendirme postas� gelir. Bu ayar� her mesaj yaz���n�zda de�i�tirebilirsiniz."
Const strTxtNotifyMeOfPrivateMessages = "�zel mesaj ald���mda eposta yoluyla beni bilgilendir"
Const strTxtAlwaysAttachMySignature = "Mesajlar�mda herzaman imza kullan"
Const strTxtEnableTheWindowsIEWYSIWYGPostEditor = "WYSIWYG edit�r�n� etkinle�tir <br /><span class=""smText"">Sadece yeni nesil taray�c�larda bu �zellik kullan�labilinir, taray�c�n�z taraf�ndan edit�r otomatik olarak alg�lan�r.</span>"
Const strTxtTimezone = "Forum saatine g�re saat dilimi"
Const strTxtPresentServerTimeIs = "�u anda sunucudaki tarih ve saat: "
Const strTxtDateFormat = "Tarih Format�"
Const strTxtDayMonthYear = "G�n/Ay/Y�l"
Const strTxtMonthDayYear = "Ay/G�n/Y�l"
Const strTxtYearMonthDay = "Y�l/Ay/G�n"
Const strTxtYearDayMonth = "Y�l/G�n/Ay"
Const strTxtHours = "saatler"
Const strTxtDay = "G�n"
Const strTxtCMonth = "Ay"
Const strTxtCYear = "Y�l"
Const strTxtRealName = "Ger�ek isim"
Const strTxtMemberTitle = "�ye ba�l���"


'Polls
'---------------------------------------------------------------------------------
Const strTxtCreateNewPoll = "Yeni anket olu�tur"
Const strTxtPollQuestion = "Anket&nbsp;Sorusu"
Const strTxtPollChoice = "Anket ��k"
Const strTxtErrorPollQuestion = "Anket Sorusu \t- Anket i�in soru belirtiniz"
Const strTxtErrorPollChoice = "Anket ��kk� \t- Anket i�in en az iki tane ��k belirleyiniz"
Const strTxtSorryYouDoNotHavePermissionToCreatePollsForum = "Forumda anket olu�turma yetkiniz bulunmamaktad�r."
Const strTxtAllowMultipleVotes = "Bu ankette birden fazla oy vermeyi etkinle�tir."
Const strTxtMakePollOnlyNoReplies = "Sadece anket olu�tur (cevaplara izin verilmez)"
Const strTxtYourNoValidPoll = "Anketiniz ge�erli bir soruyu veya ��klar� i�ermemektedir."
Const strTxtPoll = "Anket:"
Const strTxtVote = "Oy"
Const strTxtVotes = "Oylar"
Const strTxtCastMyVote = " Oy ver"
Const strTxtPollStatistics = "Anket istatistikleri"
Const strTxtThisTopicIsClosedNoNewVotesAccepted = "Bu anket kapat�lm��t�r, yeni oylar kabul edilmemektedir"
Const strTxtYouHaveAlreadyVotedInThisPoll = "Daha �nce bu ankete oy verdiniz"
Const strTxtThankYouForCastingYourVote = "Oy verdi�iniz i�in te�ekk�r ederiz."
Const strsTxYouCanNotNotVoteInThisPoll = "Bu ankete oy veremezsiniz"
Const strTxtYouDidNotSelectAChoiceForYourVote = "Oyunuz say�lmam��t�r.\n\nOy vermeniz i�in herhangi bir ��kk� i�aretlemi� olman�z laz�m."
Const strTxtThisIsAPollOnlyYouCanNotReply = "Sadece anket i�indir, mesaj yollayamazs�n�z."


'Email Notify
'---------------------------------------------------------------------------------
Const strTxtWatchThisTopic = "Bu konuyu takip et"
Const strTxtUn = "Un-"
Const strTxtWatchThisForum = "Bu forumu takip et"
Const strTxtYouAreNowBeNotifiedOfPostsInThisForum = "Bu forumdaki t�m mesajlar i�in eposta yoluyla bilgilendirme alacaks�n�z.\n\nBilgilendirme postalar�n� istemiyorsan�z \'Forumu takip etme\' \n butonuna t�klay�n�z veya Forum Se�enekleri sayfas�ndaki \'EPosta Bilgilendirme\' sayfas�n� ziyaret ediniz."
Const strTxtYouAreNowNOTBeNotifiedOfPostsInThisForum = "Bu forumdaki t�m mesajlar i�in art�k eposta almayacaks�n�z.\n\nBilgilendirme postalar�n� istiyorsan�z \'Forumu Takip Et\' \n butonuna t�klay�n�z veya Forum Se�enekleri sayfas�ndaki \'EPosta Bilgilendirme\' sayfas�n� ziyaret ediniz."
Const strTxtYouWillNowBeNotifiedOfAllReplies = "Bu konudaki mesaj�n�za g�nderilen t�m cevaplar i�in bilgilendirme postalar� alacaks�n�z.\n\nBilgilendirme postalar�n� istemiyorsan�z \'Konuyu takip etme\' \n butonuna t�klay�n�z veya Forum Se�enekleri sayfas�ndaki \'EPosta Bilgilendirme\' sayfas�n� ziyaret ediniz."
Const strTxtYouWillNowNOTBeNotifiedOfAllReplies = "Bu konudaki mesaj�n�za g�nderilen t�m cevaplar i�in art�k eposta almayacaks�n�z.\n\nBilgilendirme postalar�n� istiyorsan�z \'Konuyu Takip Et\' butonuna t�klay�n�z."


'email_messenger.asp
'---------------------------------------------------------------------------------
Const strTxtEmailMessenger = "Email Messenger"
Const strTxtRecipient = "Al�c�"
Const strTxtNoHTMLorForumCodeInEmailBody = "G�nderece�iniz eposta sadece metin tabanl�d�r (HTML kodlar� veya Forum kodlar� kullan�lamaz).<br /><br />Cevaplama adresi eposta adresiniz olacakt�r."
Const strTxtYourEmailHasBeenSentTo = "EPostan�z g�nderildi"
Const strTxtYouCanNotEmail = "Eposta g�nderemezsiniz"
Const strTxtYouDontHaveAValidEmailAddr = "Profilinizde ge�erli bir eposta adresi bulunmamaktad�r."
Const strTxtTheyHaveChoosenToHideThierEmailAddr = "se�ilen �yeler eposta adreslerini gizlemi�ler."
Const strTxtTheyDontHaveAValidEmailAddr = "se�ilen �yelerin profillerinde ge�erli bir eposta adresi bulunmamaktad�r."
Const strTxtSendACopyOfThisEmailToMyself = "G�nderilen epostan�n bir kopyas�n� kendime g�nder"
Const strTxtTheFollowingEmailHasBeenSentToYouBy = "A�a��daki epostay� size g�nderen"
Const strTxtFromYourAccountOnThe = "from the forum your participate in on "
Const strTxtIfThisMessageIsAbusive = "E�er gelen eposta spam ise veya rahats�zl�k verici ise l�tfen webmaster ile veya forum yetkilileri ile ileti�im kurunuz"
Const strTxtIncludeThisEmailAndTheFollowing = "Bu epostay� ve devam�n� ekle"
Const strTxtReplyToEmailSetTo = "L�tfen bu epostan�n yan�tlama/cevap adresini belirtiniz"
Const strTxtMessageSent = "Posta g�nderildi"



'forum_closed.asp
'---------------------------------------------------------------------------------
Const strTxtForumClosed = "Forum Kapal�"
Const strTxtSorryTheForumsAreClosedForMaintenance = "Bak�m �al��malar� sebebiyle forum kapal�d�r.<br />L�tfen daha sonra deneyiniz."


'report_post.asp
'---------------------------------------------------------------------------------
Const strTxtReportPost = "Mesaj� bildir"
Const strTxtSendReport = "Raporu g�nder"
Const strTxtProblemWithPost = "Mesajdaki problem"
Const strTxtPleaseStateProblemWithPost = "L�tfen bu mesajla ilgili s�k�nt�n�z� yaz�n�z, mesaj�n bir kopyas� forum adminlerine ve modarat�rlerine g�nderilecektir."
Const strTxtTheFollowingReportSubmittedBy = "A�a��daki raporu g�nderen"
Const strTxtWhoHasTheFollowingIssue = "who has the following issue with this post"
Const strTxtToViewThePostClickTheLink = "Mesaj� g�r�nt�lemek i�in a�a��daki linke t�klay�n�z"
Const strTxtIssueWithPostOn = "Issue With Post on"
Const strTxtYourReportEmailHasBeenSent = "Epostan�z Forum Adminlerine ve Modarat�rlerine g�nderilmi�tir."


'New from version 7.5
'---------------------------------------------------------------------------------
Const strTxtQuickLogin = "H�zl� Giri�"
Const strTxtThisTopicWasStarted = "Konu a��lma tarihi: "
Const strTxtResendActivationEmail = "Aktivasyon epostas�n� tekrar g�nder"
Const strTxtNoOfStars = "Y�ld�z say�s�"
Const strTxtOnLine2 = "Aktif"
Const strTxtCode = "Kod"
Const strTxtCodeandFixedWidthData = "Kod ve sabit geni�lik datas�"
Const strTxtQuoting = "Al�nt�"
Const strTxtMyCodeData = "Kodum ve sabit geni�lik datas�"
Const strTxtQuotedMessage = "Al�nt� yap�lm�� mesaj"
Const strTxtWithUsername = "Kullan�c� ad�yla birlikte"
Const strTxtGo = "Git"
Const strTxtDataBasedOnActiveUsersInTheLastXMinutes = "Bu bilgiler son 20 dakika i�inde aktif olan �yeleri kapsar"
Const strTxtSoftwareVersion = "Yaz�l�m Versiyonu"
Const strTxtForumMembershipNotAct = "Forum �yeli�iniz hen�z aktive edilmemi�!"
Const strTxtMustBeRegisteredToPost = "Mesajlarda s�ralama yapabilmeniz i�in forum �yesi olman�z gerekmektedir."
Const strTxtMemberCPMenu = "�ye Kontrol Paneli"
Const strTxtYouCanAccessCP = "Forum ara�lar� ve Forum Se�eneklerini de�i�tirebilirsiniz "
Const strTxtEditMembersSettings = "Bu �yenin forum se�eneklerini de�i�tir"
Const strTxtSecurityCodeConfirmation = "G�venlik Kodu Onay� (gerekli)"
Const strTxtUniqueSecurityCode = "G�venlik Kodu"
Const strTxtEnterCAPTCHAcode = "L�tfen resimde g�rd���n�z kodu G�venlik Kodu alan�na giriniz.<br />Taray�c�n�z�n �erez deste�inin a��k olmas� gerekmektedir."
Const strTxtErrorSecurityCode = "G�venlik Kodu \t- Resimde g�rd���n�z kodu girmelisiniz"
Const strTxtSecurityCodeDidNotMatch = "Girmi� oldu�unuz g�venlik kodu ile resimdeki kod uyu�mamaktad�r.\n\nYeni bir g�venlik kodu resmi olu�turulmu�tur."

'login_user_test.asp
'---------------------------------------------------------------------------------
Const strTxtSuccessfulLogin = "Giri� ba�ar�l�"
Const strTxtSuccessfulLoginReturnToForum = "Ba�ar�yla giri� yapt�n�z, l�tfen bekleyiniz foruma y�nlendiriliyorsunuz"
Const strTxtUnSuccessfulLoginText = "�erez sorunundan dolay� giri�iniz ba�ar�s�z olmu�tur. <br /><br />L�tfen taray�c�n�z�n �erez deste�inin a��k oldu�undan ve IP Adresinizin gizli olmad���ndan emin olunuz."
Const strTxtUnSuccessfulLoginReTry = "Buraya t�klayarak foruma giri�i tekrar deneyebilirsiniz."
Const strTxtToActivateYourForumMem = "Forum �yeli�inizin aktif olmas� i�in kay�t olduktan sonra eposta adresinize gelen linke t�klaman�z gerekmektedir."

'email_notify_subscriptions.asp
'---------------------------------------------------------------------------------
Const strTxtEmailNotificationSubscriptions = "EPosta Bilgilendirme"
Const strTxtSelectForumErrorMsg = "Forum Se�iniz\t- Bilgilendirme postalar�n� istedi�iniz forumu se�iniz"
Const strTxtYouHaveNoSubToEmailNotify = "EPosta ile bilgilendirme talimat�n�z bulunmamaktad�r"
Const strTxtThatYouHaveSubscribedTo = "EPosta bilgilendirme talimatlar�n�zd�r a�a��dad�r"
Const strTxtUnsusbribe = "Takip etme"
Const strTxtAreYouWantToUnsubscribe = "Bunlar�n takip edilmemesini istedi�inizden emin misiniz?"



'New from version 7.51
'---------------------------------------------------------------------------------
Const strTxtSubscribeToForum = "Yeni mesajlar� takip et. (EPosta Bilgilendirme ile)"
Const strTxtSelectForumToSubscribeTo = "Takip etmek istedi�iniz forumu se�iniz"


'New from version 7.7
'---------------------------------------------------------------------------------
Const strTxtOnlineStatus = "Online"
Const strTxtOffLine = "Offline"


'New from version 7.8
'---------------------------------------------------------------------------------
Const strTxtConfirmOldPass = "Eski �ifreyi Onayla"
Const strTxtConformOldPassNotMatching = "�ifre do�rulamas� kay�tlar�m�zdaki tan�mlaman�z ile uyu�muyor.\n\nE�er �ifrenizi de�i�tirmek istiyorsan�z l�tfen eski �ifrenizi do�ru giriniz"



'New from version 8.0
'---------------------------------------------------------------------------------
Const strTxtSub = "Alt"
Const strTxtHidden = "Gizli"
Const strTxtHidePost = "Mesaj� Gizle"
Const strTxtAreYouSureYouWantToHidePost = "Bu mesaj� gizlemek istedi�inizden emin misiniz?"
Const strTxtModeratedPost = "Pre-Approved Post"
Const strTxtYouArePostingModeratedForum = "You are posting in a moderated forum."
Const strTxtBeforePostDisplayedAuthorised = "Mesaj�n�z�n forumda yay�nlanabilmesi i�in forum adminleri ve moderat�rler taraf�ndan onaylanmas� gerekmektedir."
Const strTxtHiddenTopics = "Moderated Topics"
Const strTxtVerifiedBy = "Onaylayan"
Const strTxtYourEmailHasChanged = "EPosta adresiniz"
Const strTxtPleaseUseLinkToReactivate = "olarak de�i�tirildi, l�tfen �yeli�inizin tekrar aktivasyonu i�in linke t�klay�n�z"
Const strTxtToday = "Bug�n"
Const strTxtPreviewPost = "Mesaj �nizleme"
Const strTxtEmailNotify = "Cevap geldi�inde EPosta ile bilgilendir"
Const strTxtAvatarUpload = "Avatar y�kle"
Const strTxtClickOnEmoticonToAdd = "Mesaj�n�za eklemek istedi�iniz emoticon'a t�klay�n�z."
Const strTxtUpdatePost = "Mesaj� G�ncelle"
Const strTxtShowSignature = "�mzam� G�ster"
Const strTxtQuickReply = "H�zl� Cevap"
Const strTxtCategory = "Kategori"
Const strTxtReverseSortOrder = "Tersinden S�rala"
Const strTxtSendPM = "�zel Mesaj G�nder"
Const strTxtSearchKeywords = "Anahtar s�zc�kleri ara"
Const strTxtSearchbyKeyWord = "Anahtar s�zc�klere g�re ara"
Const strTxtSearchbyUserName = "Kullan�c� ad�na g�re ara (�ste�e Ba�l�)"
Const strTxtMatch = "E�le�en"
Const strTxtSearchOptions = "Arama Ayarlar�"
Const strTxtCtrlApple = "('control' veya 'apple' tu�una basarak birden fazla se�ebilirisniz)"
Const strTxtFindPosts = "Mesajlarda ara"
Const strTxtAndNewer = "ve Yeniler"
Const strTxtAndOlder = "ve Eskiler"
Const strTxtAnyDate = "Herhangi bir zaman"
Const strTxtNumberReplies = "Cevaplanma Say�s�na G�re"
Const strTxtExactMatch = "Tam E�le�en"
Const strTxtSearhExpiredOrNoPermission = "Bu arama ge�erli de�il veya arama yapmaya yetkiniz bulunmamaktad�r"
Const strTxtCreateNewSearch = "Yeni arama olu�tur"
Const strTxtNoSearchResultsFound = "Hi� sonu� bulunamad�"
Const strTxtSearchError = "Arama Hatas�"
Const strTxtSearchWordLengthError = "Araman�zda 3 karakterden az kelime veya kelimeler var"
Const strTxtIPSearchError = "IP Adresinize izin verilen arama limitini a�t�n�z<br /><br />L�tfen yeni arama yapmadan �nce 30sn bekleyiniz"
Const strTxtResultsIn = "Sonu�lar"
Const strTxtSecounds = "saniyede olu�turuldu"
Const strTxtFor = "i�in"
Const strTxtThisSearchWasProcessed = "This search was processed"
Const strTxtError = "Hata"
Const strTxtReply = "Cevap"
Const strTxtClose = "Kapat"
Const strTxtActiveStats = "Active Stats"
Const strTxtInformation = "Bilgilendirme"
Const strTxtCommunicate = "�leti�im"
Const strTxtDisplayResultsAs = "Sonu�lar� �u �ekilde g�ster"
Const strTxtViewPost = "Mesaj� g�ster"
Const strTxtPasswordRequiredViewPost = "Mesaj� g�r�nt�lemek i�in �ifre gerekli"
Const strTxtNewestPostFirst = "Yeni mesajlar ba�ta"
Const strTxtOldestPostFirst = "Eski mesajlar ba�ta"
Const strTxtMessageIcon = "Mesaj ikonu"
Const strTxtSkypeName = "Skype Ad�"
Const strTxtLastPostDetailNotHiddenDetails = "Please note:- Last Post details don't include details of hidden posts."
Const strTxtOriginallyPostedBy = "Orjinalini yazan:"
Const strTxtViewingTopic = "G�r�nt�ledi�i Konu:"
Const strTxtViewingIndex = "Giri�i G�r�nt�l�yor"
Const strTxtForumIndex = "Forum Giri�i"
Const strTxtIndex = "Giri�"
Const strTxtViewing = "Ki�i G�r�nt�l�yor"
Const strTxtSearchingForums = "Forumlar� Ar�yor"
Const strTxtSearchingFor = "Bunu Ar�yor"
Const strTxtWritingPrivateMessage = "�zel Mesaj Yaz�yor"
Const strTxtViewingPrivateMessage = "�zel Mesaj G�r�nt�l�yor"
Const strTxtEditingPost = "Mesaj D�zenliyor"
Const strTxtWritingReply = "Cevap Yaz�yor"
Const strTxtWritingNewPost = "Yeni Mesaj Yaz�yor"
Const strTxtCreatingNewPost = "Yeni Anket Olu�turuyor"
Const strTxtWhatsGoingOn = "Forumda Neler Oluyor?"
Const strTxtLoadNewCode = "Yeni Kod Y�kle"
Const strTxtApprovePost = "Mesaj� Onayla"
Const strTxt3LoginAtteptsMade = "Bu kullan�c� i�in 3 giri� denemesi yap�lm��t�r.<br />L�tfen bilgilerinizi girdikten sonra g�venlik kodunuda girin."
Const strTxtSuspendUser = "�yeli�i Ask�ya Al"
Const strTxtAdminNotes = "Y�netici/Moderat�r Notu"
Const strTxtAdminNotesAbout = "Bu b�l�me yazaca��n�z notu sadece y�neticiler ve moderat�rler ki�inin profiline bakt�klar�nda g�rebilir. �ye hakk�nda uyar�lar v.b. yazabilirsiniz(max 250 karakter)"
Const strTxtAge = "Ya�"
Const strTxtUnknown = "Ge�ersiz"
Const strTxtSuspended = "Ask�ya Al�nd�"
Const strTxtEmailNewUserRegistered = "A�a��da yeni kaydolan �yeler listelenmektedir "
Const strTxtToActivateTheNewMembershipFor = "Yeni �yeli�i aktifle�tirmek i�in "
Const strTxtNewMemberActivation = "Yeni �ye Aktivasyonu"
Const strTxtEmailYouCanNowUseOnceYourAccountIsActivatedTheForumAt = "Giri� bilgileriniz a�a��dad�r. �yeli�iniz Forum Y�neticisi taraf�ndan onayland�ktan sonra yeni mesaj g�nderebilir, mesajlar� cevaplayabilirsiniz"
Const strTxtYouAdminNeedsToActivateYourMembership = "<strong>�yeli�inizin Forum Y�neticisi taraf�ndan onaylanmas� gerekmektedir!</strong>"
Const strTxtEmailYourForumMembershipIsActivatedThe = "Forum �yeli�iniz �u anda aktifle�tirilmi�tir.Yeni mesaj yazabilir, mesajlara cevap verebilirsiniz."
Const strTxtTheAccountIsNowActive = "Hesap aktifle�tirildi!!"
Const strTxtErrorOccuredActivatingTheAccount = "Hesab�n aktifle�tirilmesi s�ras�nda bir sorunla kar��la��ld�"
Const strTxtMustBeLoggedInAsAdminActivateAccount = "Yeni �yelerin aktivasyonunu yapabilmek i�in y�netici olarak giri� yapm�� olman�z gerekmektedir. <br /> Y�netici giri�i yapt�ktan sonra e-postadaki linki tekrar t�klay�n."
Const strTxtTodaysBirthdays = "Bug�n Do�um G�n� Olan �yeler"
Const strTxtCalendar = "Takvim"
Const strTxtEventDate = "Olay�n Tarihi"
Const strTxtEvent = "Olay"
Const strTxtCalendarEvent = "Takvim Olay�"
Const strTxtLast = "Son"
Const strTxtRSS = "RSS"
Const strTxtNewPostFeed = "Yeni Mesaj Linki"
Const strTxtLastTwoDays = "Son 2 G�n"
Const strTxtThisRSSFileIntendedToBeSyndicated = "Bu sayfa RSS okuyucular ve web sayfalar�nda e� zamanl� g�sterim i�in tasarlanm��t�r."
Const strTxtCurrentFeedContent = "G�ncel link i�eri�i"
Const strTxtSyndicatedForumContent = "G�ncel forum i�erigi"
Const strTxtSubscribeNow = "RSS Linkini Al!"
Const strTxtSubscribeWithWebBasedRSS = "se�iminizi t�klay�n"
Const strTxtWithOtherReaders = "e�er bilgisayar�n�zda RSS okuyucu y�kl� ise "
Const strTxtSelectYourReader = "Okuyucunuzu Se�in"
Const strTxtThisIsAnXMLFeedOf = "XML i�erik linki"
Const strTxtDirectLinkToThisPost = "Mesaj�n Direkt Linki"
Const strTxtWhatIsAnRSSFeed = "RSS Linki Nedir?"


'New from version 8.02
'---------------------------------------------------------------------------------
Const strTxtSecurityCodeDidNotMatch2 = "Girdi�iniz g�venlik kodu ekranda g�sterilen ile ayn� de�il."


'New from version 8.05
'---------------------------------------------------------------------------------
Const strTxtPleaseDontForgetYourPassword = "L�tfen �ifrenizi unutmay�n�z, �ifre veritaban�nda kodlanarak sakland��� i�in unuttu�unuz �ifreyi geri alma olana�� yoktur. Unutman�z durumunda Parolam� Unuttum b�l�m�nden Kullan�c� Ad�n�z� ve E-posta adresinizi belirterek, yeni bir �ifrenin E-posta adresinize g�nderilmesini isteyebilirsiniz."
Const strTxtActivationEmail = "Aktivasyon E-postas�" 
Const strTxtTopicReplyNotification = "Konu Cevap Bildirimi"
Const strTxtUserNameOrEmailAddress = "Kullan�c� Ad� veya E-posta Adresi"
Const strTxtAnonymousMembers = "Bilinmeyen �ye"
Const strTxtGuests = "Misafir"
Const strTxtNewPosts = "yeni mesajlar"
Const strTxtNoNewPosts = "yeni mesaj yok"
Const strTxtFullReplyEditor = "Tam Edit�r"


'New from version 9
'---------------------------------------------------------------------------------
Const strTxtForumHome = "Anasayfa"
Const strTxtNewMessages = "Yeni Mesaj"
Const strTxtsoh = "Chat (12 Online)"
Const strTxtFAQ = "Yard�m"
Const strTxtsohbet = "Sohbet"
Const strTxtUnAnsweredTopics = "Cevaplanmam�� Konular"
Const strTxtShowPosts = "Mesajlar� G�ster"
Const strTxtModeratorTools = "Moderat�r Ara�lar�"
Const strTxtResyncTopicPostCount = "Forum �statistiklerini G�ncelle"
Const strTxtAdminControlPanel = "Y�netici Kontrol Paneli"
Const strTxtAdvancedSearch = "Geli�mi� Arama"
Const strTxtLockTopic = "Konuyu Kilitle"
Const strTxtHideTopic = "Konuyu Gizle"
Const strTxtShowTopic = "Konuyu G�ster"
Const strTxtTopicOptions = "Konu Se�enekleri"
Const strTxtForumOptions = "Forum Se�enekleri"
Const strTxtFindMembersPosts = "�yenin Mesajlar�n� Bul"
Const strTxtMembersProfile = "�ye Profili"
Const strTxtVisitMembersHomepage = "�yenin Web Sitesine Git"
Const strTxtFirstPage = "�lk Sayfa"
Const strTxtLastPage = "Son Sayfa"
Const strTxtPostOptions = "Mesaj Se�enekleri"
Const strTxtBlockUsersIP = "IP Engelle"
Const strTxtCreateNewTopic = "Yeni Konu Olu�tur"
Const strTxtNewPoll = "Anket"
Const strTxtControlPanel = "Kontrol Paneli"
Const strTxtSubscriptions = "Abonelikler"
Const strTxtMessenger = "Haberci"
Const strTxtBuddyList = "Arkada� Listesi"
Const strTxtProfile2 = "Profil"
Const strTxtSubscribe = "Abone Ol"
Const strTxtMultiplePages = "Bir�ok Sayfa"
Const strTxtCurrentPage = "Ge�erli Sayfa"
Const strTxtRefreshPage = "Sayfay� Yenile"
Const strTxtAnnouncements = "Duyurular"
Const strTxtHiddenTopic = "D�zenlenmi� Konu"
Const strTxtHot = "S�cak"
Const strTxtLocked = "Kilitli"
Const strTxtNewPost = "Yeni Mesaj"
Const strTxtPoll2 = "Anket"
Const strTxtSticky = "Sabit"
Const strTxtForumPermissions = "Forum �zinleri"
Const strTxtForumWithSubForums = "Forum ile Alt Forum"
Const strTxtPostNewTopic2 = "Yeni Konu A�"
Const strTxtViewDropDown = "A��l�r Kutu G�r"
Const strTxtFull = "dolu"
Const strNotYetRegistered = "�ye ol"
Const strTxtNewsletterSubscription = "Haber Aboneli�i"
Const strTxtSignupToRecieveNewsletters = "Haberleri almak i�in �ye ol " 
Const strTxtNewsBulletins = "Haber B�ltenleri"
Const strTxtPublished = "Yay�mland�"
Const strTxtStartDate = "Ba�lang�� Tarihi"
Const strTxtEndDate = "Biti� Tarihi"
Const strTxtNotRequiredForSingleDateEvents = "not required for single date events"
Const strTxtIn = ""
Const strTxtGender = "Cinsiyet"
Const strTxtMale = "Bay"
Const strTxtFemale = "Bayan"



Const strTxtFileAlreadyUploaded = "Ayn� isimde dosyay� daha �nceden y�klemi�siniz!"
Const strTxtSelectOrRenameFile = "L�tfen ba�ka bir dosya se�iniz veya dosyan�n ad�n� de�i�tirip tekrar deneyin."
Const strTxtAllotedFileSpaceExceeded = "Size ayr�lan dosya alan�n� a�t�n�z: "
Const strTxtDeleteFileOrImagesUingCP = "L�tfen �ye Kontrol Panelinizdeki Dosya Y�netimini kullanarak kullanmad���n�z dosya ve resimleri silin."



'File Manager
Const strTxtFileManager = "Dosya Y�netimi"
Const strTxtFileName = "Dosya Ad�"
Const strTxtFileSize = "Dosya Boyutu"
Const strTxtFileType = "Dosya T�r�"
Const strTxtFileExplorer = "Dosyalar"
Const strTxtFileProperties = "Dosya �zellikleri"
Const strTxtFilePreview = "Dosya �nizleme"
Const strTxtAllocatedFileSpace = "Dosya Alan� �statistikleri"
Const strTxtYouHaveUsed  = "Kulland���n�z Alan:"
Const strTxtFromYour = "&nbsp;&nbsp;&nbsp;�zin Verilen Alan:"
Const strTxtOfAllocatedFileSpace = ""
Const strTxtYourFileSpaceIs = "Dosya alan�n�z"
Const strTxtDownloadFile = "Y�klenen Dosya"
Const strTxtNewUpload = "Yeni Y�kleme Yap"
Const strTxtDeleteFile = "Dosya Sil"
Const strTxtRenameFile = "Yeniden Adland�r"
Const strTxtAreYouSureDeleteFile = "Bu dosyay� silmek istedi�inize emin misiniz?"
Const strTxtNoFileSelected = "Se�ili dosya yok"
Const strTxtTheFileNowDeleted = "Dosya silindi"
Const strTxtYourFileHasBeenSuccessfullyUploaded = "Dosyan�z ba�ar�l� bir �ekilde y�klendi."
Const strTxtSelectUploadType = "Y�klem T�r� Se�"
Const strTxtYouTube = "YouTube"
Const strTxtUploadFolderEmpty = "Y�kleme Klas�r� Bo�"




'New from version 9.04
'---------------------------------------------------------------------------------

Const strTxtAutologinOnlyAppliesToSession = ""
Const strTxtViewUnreadPost = "Okunmam�� Mesajlar� G�r"



'New from version 9.51
'---------------------------------------------------------------------------------

Const strTxtPendingApproval = "Onay bekliyor"
Const strTxtThatRequiresApproval = "bunlar onay bekliyor."

Const strTxtMovieProperties = "Movie �zellikleri"
Const strTxtMovieType = "Movie Tipi"
Const strTxtYouTubeFileName = "YouTube Dosya ad�"
Const strTxtFlashMovieURL = "Flash Movie URL"



'New from version 9.52
'---------------------------------------------------------------------------------
Const strTxtThroughTheirForumProfileAtLinkBelow = "through their forum profile at the link below."
Const strTxtYouCanNotEmailTisTopicToAFriend = "You can not email this topic to a friend"
Const strTxtToReplyPleaseEmailContact = "To reply to this email contact"
Const strTxtInsertMovie = "Flash Movie Ekle"


'New from version 9.54
'---------------------------------------------------------------------------------
Const strTxtTheEmailFailedToSendPleaseContactAdmin = "EPosta g�nderimi ba�ar�s�z oldu. L�tfen hata mesaj�yla birlikte forum adminlerine ula�on."
Const strTxtFindMember = "�ye bul"
Const strTxtSearchForTopicsThisMemberStarted = "Bu �yenin a�m�� oldu�u konular� bul"
Const strTxtMemberName2 = "�ye Ad�"
Const strTxtSearchTimeoutPleaseNarrowSearchTryAgain = "Araman�z zaman a��m�na u�rad�. L�tfen arama kriterlerinizi g�zden ge�irip tekrar deneyin"


'New from version 9.55
'---------------------------------------------------------------------------------
Const strTxtTheFileFailedTHeSecurityuScanAndHasBeenDeleted = "Dosya g�venlik taramas�n� ge�emedi ve silindi, dosya i�inde zararl� kodlar olabilir."


'New from version 9.56
'---------------------------------------------------------------------------------
Const strTxtShareTopic = "Konuyu Payla�"
Const strTxtPostThisTopicTo = "Bu kadar konuyu ilan et"

'New from version 9.61
'---------------------------------------------------------------------------------
Const strTxtSponsor = "Sponsorlar"

'New from version 9.64
'---------------------------------------------------------------------------------
Const strTxtIPAddress = "IP Adresi"


'New from version 9.65
'---------------------------------------------------------------------------------
Const strTxtTranslate = "�eviri"

'New from version 9.66
'---------------------------------------------------------------------------------
Const strTxtConfirmEmail = "E-Posta Onayla"
Const strTxtErrorConfirmEmail = "E-posta Onaylama ��in Alanlar E�le�miyor"


'New from version 9.67
'---------------------------------------------------------------------------------
Const strTxtThereMayAlsoBeOtherMessagesPostedOn = "There may also be other messages posted on"
Const strTxtWarningYourSessionHasExpiredRefreshPageFormDataWillBeLost = "UYARI\nOturumunuz zaman a��m�na u�rad�! Taray�c�n�zdaki 'Yenile\' butonuna basarak sayfay� yenileyin.\n** Girmi� oldu�unuz form datalar� kaybolacakt�r! **"


'New from version 9.70
'---------------------------------------------------------------------------------
Const strTxtNoFollowAppliedToAllLinks = "NoFollow forumdaki t�m linkler i�in aktif hale getirildi (rel=""nofollow"")"

'New from version 9.71
'---------------------------------------------------------------------------------
Const strTxtViewIn = "G�r�n�m"
Const strTxtMoble = "Mobil"
Const strTxtClassic = "Klasik"

'New from version 10
'---------------------------------------------------------------------------------
Const strTxtStatus = "Durum"
Const strTxtTheEmailAddressEnteredIsInvalid = "E-Mail adresi hatal�"
Const strTxtMostUsersEverOnlineWas = "Bug�ne kadar en fazla online olan ki�i say�s�"
Const strTxtTypeTheNameOfMemberInBoxBelow = "Bulmak istedi�iniz �ye ad�n�n tamam�n� veya bir k�sm�n� yaz�n.."
Const strTxtSelectNameOfMemberFromDropDownBelow = "Below is a list of members who match your search criteria, select the member you are looking for and click the 'Select' button to insert this members name into the form this box opened from."
Const strTxtCharacters = "karakterler"
Const strTxtMxLFailedLoginAttemptsMade = "More than the maximum failed login attempts have been made on this account.<br />Please enter your details again, including security code."
Const strTxtNumberOfPoints = "Noktalar�n Say�s�"
Const strTxtPoints = "Puanlar"
Const strTxtPasswordNotComplex = "�ifreniz kar���k karakterlerden olu�mal�d�r.\nEn az 1 b�y�k harf, 1 k���k harf ve 1 say� i�ermelidir."
Const strTxtLadderGroup = "Merdiven Grup"
Const strTxtNone = "Hi�biri"
Const strTxtRealNameError = "Ger�ek ad�n�z� giriniz"
Const strTxtLocationError = "Lokasyon giriniz"
Const strTxtNotRequired = "Gerekli de�il"
Const strTxtRating = "Oylama"
Const strTxtTopicRating = "Ba�l�k oylama"
Const strTxtAverage = "Ortalama"
Const strTxtRateTopic = "Konu Oran�"
Const strTxtYouHaveAlreadyRatedThisTopic = "Bu konu i�in zaten oy kullanm��t�n�z!"
Const strTxtYouCanNotRateATopicYouStarted = "You can not rate a Topic that you started!"
Const strTxtThankYouForRatingThisTopic = "Oy kulland�n�z i�in te�ekk�rler."
Const strTxtExcellent = "M�kemmel"
Const strTxtPoor = "K�t�"
Const strTxtGood = "Iyi"
Const strTxtTerrible = "Korkun�"
Const strTxtRateThisTopicAs = "Konu oran� i�in buray� t�klay�n"
Const strTxtSortOrder = "S�rala"
Const strTxtMost = "En"
Const strTxtHighestRating = "Y�ksek oranl�"
Const strTxtBy1 = "taraf�ndan"
Const strTxtMembers2 = "�yeler"
Const strTxtEvents = "Etkinlikler"
Const strTxtYouAreOnlyPermittedToEditPostWithin = "Mesaj� d�zenleyebilirsiniz"
Const strTxtChat = "Sohbet"
Const strTxtOnlineMembers = "Online Kullan�c�lar"
Const strTxtReason = "Ak�l"
Const strTxtYourMessageWasRejectedByTheSpamFilters = "Mesaj�n�z spam filtreleri taraf�ndan reddedildi"
Const strYouMustEnterYour = "Girmelisiniz"
Const strTxtViewUnreadPost1 = "Okunmam�� Mesajlar"
Const strTxtSetAsAnswer = "Cevap olarak ayarla"
Const strTxtUnSetAsAnswer = "Cevap olarak ayarlamaktan kald�r"
Const strTxtExternalLinkTo = "D��ar�ya ba�lant� adresi"
Const strTxtYouMustBeARegisteredMemberAndPostAReplyToViewMessage = "Gizli i�eri�i g�rmek i�in kay�t olup ve konuya mesaj yazman�z gerekmektedir."
Const strTxtHideContent = "��eri�i Gizle"
Const strTxtPostContentHiddenUntilReply = "Gizli i�eri�i g�rmek i�in konuya cevap yazman�k gerekir."


Const strTxtThanks = "Te�ekk�rler"
Const strTxtThanked = "Te�ekk�r Edildi"
Const strTxtYouMustHaveAnActiveMemberAccount  = "Aktif �ye Hesab� olmal�d�r"
Const strTxtYouCanNotThankYourself = "Kendinize te�ekk�r edemezsiniz."
Const strTxtYouHaveAlreadySaidThanksForThisPost = "Zaten mesaj i�in te�ekk�r edildi."
Const strTxtHasBeenThankedForTheirPost = "has been thanked for their Post"
Const strTxtYourPmInboxIsFullPleaseDeleteOldPMs = "�zel Mesaj Gelen Kutusu dolu! Herhangi bir eski veya istenmeyen �zel mesajlar�n�z� silin."
Const strTxtShare = "Payla�"
Const strTxtShareThisPageOnTheseSites ="Sayfay� Payla�"

Const strTxtFacebook = "Facebook"
Const strTxtLinkedIn = "LinkedIn"
Const strTxtTwitter = "Twitter"

Const strTxtAnswer = "Cevap"
Const strTxtResolution = "��z�m"
Const strTxtOfficialResponse = "Resmi Yan�t"
%>