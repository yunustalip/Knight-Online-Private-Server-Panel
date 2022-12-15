<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2009 Web Wiz(TM). All Rights Reserved.
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


'pm_welcome.asp
'---------------------------------------------------------------------------------
Const strTxtToYourPrivateMessenger = "�zel Mesajlara"
Const strTxtPmIntroduction = "�zel Mesaj Sistemi �le Di�er Forum �yeleri �le Mesajla�abilirsiniz."
Const strTxtInboxStatus = "Gelen Kutusu �zellikler"
Const strTxtGoToYourInbox = "Gelen Kutusuna Git"
Const strTxtNoNewMsgsInYourInbox = "Gelen Kutunuzda Yeni Mesaj Yok."
Const strTxtYourLatestPrivateMessageIsFrom = "Mesaj�n�z Var G�nderen:"
Const strTxtSentOn = "G�nderim Zaman�:"
Const strTxtPrivateMessengerOverview = "�zel Mesaj Sistem �zellikleri"
Const strTxtInboxOverview = "Gelen Mesajlar"
Const strTxtOutboxOverview = "G�nderilmi� Mesajlar."
Const strTxtBuddyListOverview = "K�saca Adres Defteriniz."
Const strTxtNewMsgOverview = "Yeni Bir Mesaj G�ndermek ��in."

'pm_inbox.asp
'---------------------------------------------------------------------------------

Const strTxtInbox = "Gelen Kutusu"
Const strTxtNewPrivateMessage = "Yeni �zel Mesaj"
Const strTxtNoPrivateMessages = "Yeni Mesaj Yok"
Const strTxtRead = "Oku"
Const strTxtMessageTitle = "Mesaj Ba�l���"
Const strTxtMessageFrom = "G�nderen"
Const strTxtDate = "Al�nma Tarihi"
Const strTxtBlock = "Engelle"
Const strTxtSentBy = "G�nderen"
Const strTxtDeletePrivateMessageAlert = "�zel Mesaj� Silmek �stiyormusunuz ?"
Const strTxtPrivateMessagesYouCanReceiveAnother = "�zel Mesaj, Kalan Mesaj"
Const strTxtOutOf = ">"
Const strTxtPreviousPrivateMessage = "�nceki Mesaj"
Const strTxtMeassageDeleted = "�zel Mesaj Silindi !"

'pm_message.asp
'---------------------------------------------------------------------------------
Const strTxtSorryYouDontHavePermissionsPM = "�zg�n�m, �zel Mesajlar� G�r�nt�lemek ��in Yetkili De�ilsiniz."
Const strTxtYouDoNotHavePermissionViewPM = "�zel Mesajlar� G�r�nt�lemek ��in Yetkili De�ilsiniz.."
Const strTxtNotificationReadPM = "�zel Mesaj Uyac�s�n� Oku"
Const strTxtReplyToPrivateMessage = "Bu �zel Mesaja Cevap Yaz"
Const strTxtAddToBuddy = "Arkada� Listesine Ekle"
Const strTxtThisIsToNotifyYouThat = ""
Const strTxtHasReadPM = "Adl� �yemizin Mesaj�n�z� Okudu�unu"
Const strTxtYouSentToThemOn = "Belirtmektedir"


'pm_new_message_form.asp
'---------------------------------------------------------------------------------
Const strTxtSendNewMessage = "Yeni Mesaj G�nder"
Const strTxtPostMessage = "Mesaj� G�nder"
Const strTxtEmailNotifyWhenPMIsRead = "Mesaj Okundu�u Zaman E-posta �le Beni Uyar"
Const strTxtToUsername = "Kullan�c�&nbsp;�smi"
Const strSelectFormBuddyList = "Arkada� Listesinden Se�"
Const strTxtNoPMSubjectErrorMsg = "Konu \t\t- �zel Mesaj�n Konusunu Giriniz."
Const strTxtNoToUsernameErrorMsg = "Kullan�c� �smi \t- Kullan�c� �smini Girdikten Sonra Mesaj� Yollay�n�z."
Const strTxtNoPMErrorMsg = "Mesaj \t\t- �zel Mesaj�n�z� Girip Yollay�n�z."
Const strTxtSent = "G�nderen"

'pm_new_message.asp
'---------------------------------------------------------------------------------
Const strTxtAPrivateMessageHasBeenPosted = "�zel Mesaj�n�z �ste�inize Ba�l� Olarak Yolland�."
Const strTxtClickOnLinkBelowForPM = "�zel mesaj� Okumak ��in Linke T�klay�n"
Const strTxtNotificationPM = "�zel Mesaj Uyar�c�s�"
Const strTxtTheUsernameCannotBeFound = "Girdi�iniz Kullan�c� Ad� Bulunamad�.."
Const strTxtYourPrivateMessage = "Konusu Bu Olan:"
Const strTxtHasNotBeenSent = "G�nderilemedi!"
Const strTxtAmendYourPrivateMessage = "Mesaj�n�za Geri D�n�n"
Const strTxtReturnToYourPrivateMessenger = "�zel Mesajlara Geri D�n"
Const strTxtYouAreBlockedFromSendingPMsTo = "Mesajlar�n�z �letilemedi ��nk� �smi Yaz�l� Ki�i Sizi Engellenenler Listesine Ald�:"
Const strTxtHasExceededMaxNumPPMs = "S�n�rland�r�lm�� Mesaj Say�s�n� A�t��� ��in Yollanmad�."
Const strTxtHasSentYouPM = "Size A�a��daki Ba�l�k �le Mesaj Yollad�."
Const strTxtToViewThePrivateMessage = "Mesaj� G�rmek ��in"


'pm_buddy_list.asp
'---------------------------------------------------------------------------------
Const strTxtNoBuddysInList = "Arkada� Listesinde Hi� Dostunuz Yok"
Const strTxtDeleteBuddyAlert = "Bu Arkada��n�z� listeden silmek istedi�inizden eminmisiniz?"
Const strTxtNoBuddyErrorMsg = "�ye Ad� \t- Bir Forum �yesini Arkada� Listeme Ekle"
Const strTxtBuddy = "Arkada�"
Const strTxtDescription = "Tan�m"
Const strTxtContactStatus = "Haber Durumu"
Const strTxtThisPersonCanNotMessageYou = "Bu Ki�i Size Mesaj Atamaz."
Const strTxtThisPersonCanMessageYou = "Bu Ki�i Size Mesaj Atabilir."
Const strTxtAddNewBuddyToList = "Listeye Yeni Arkada� Ekle"
Const strTxtMemberName = "�ye Ad�:"
Const strTxtAllowThisMemberTo = "Bu �yenin"
Const strTxtMessageMe = "Bana Mesaj Atmas�na �zin Ver"
Const strTxtNotMessageMe = "Bana Mesaj Atmas�na �zin Verme."
Const strTxtHasNowBeenAddedToYourBuddyList = "Arkada� Listenize Eklendi."
Const strTxtIsAlreadyInYourBuddyList = "Zaten Arkada� Listenizde."
Const strTxtUserCanNotBeFoundInDatabase = "Veritaban�m�zda Kay�d� Bulunamad�.\n\n�ye Ad�n� Do�ru Yazd���n�zdan Emin Olunuz"



Const strTxtOutbox = "Giden Kutusu"
Const strTxtMessageTo = "Mesaj� Alacak Ki�i"
Const strTxtMessagesInOutBox = "Mesajlar Kar�i Taraf Okuyup Silene Kadar Giden Kutusunda Sakl� Tutulacakt�r.."

'New from version 7.02
'---------------------------------------------------------------------------------
Const strTxtYourInboxIs = "Gelen Kutunuz"
Const strTxtEmailThisPMToMe = "Bu �zel Mesaj� Kendi E-posta Adresime Yolla"
Const strTxtEmailBelowPrivateEmailThatYouRequested = "G�nderilen �zel Mesajlar�n Bir Kopyas�n� E-posta Adresime Yolla"
Const strTxtAnEmailWithPM = "�zel Mesaj ��eren E-posta,"
Const strTxtBeenSent = "E-posta Adresinize Yolland�."
Const strTxtNotBeenSent = "E-posta Adresinize Yollanamad� L�tfen Daha Sonra Tekrarlay�n�z."
Const strTxtSelected = "Se�ili"


'New from version 8
'-----------------------------------------------------------------
Const strTxtYouAreOnlyPerToSend = "Saat ba��na izin verilen mesaj say�s�:"
Const strTxtYouHaveExceededLimit = "Saat ba��na mesaj limitini doldurdunuz"


'New from version 9
'-----------------------------------------------------------------
Const strTxtNewMessage = "Yeni Mesaj"



'New from version 10
'-----------------------------------------------------------------
Const strTxtYourOutboxIs = "Your outbox is"
Const strTxtPrivateMessagesYouCanSendAnother = "Private Messages, you can send another"
Const strTxtYouAreOnlyPerToSendAMaximum = "You are only permitted to send a maximum"
Const strTxtPMsYouHaveExceededLimit = "�zel mesaj s�n�r� a��ld�"
Const strTxtToSendFutherPMsYouWillNeedToDelete = "�zel Mesajlar g�ndermek i�in Giden Kutusunda istenmeyen �zel mesajlar� silmeniz gerekir"
%>