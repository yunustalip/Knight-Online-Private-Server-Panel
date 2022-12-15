<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
'**                                               
'**  Copyright (C)2001-2011 Web Wiz(TM). All Rights Reserved.  
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
'**  http://www.webwiz.co.uk/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************



Const strTxtTextFormat = "Yaz� Bi�imi"
Const strTxtMode = "Kip (Mode)"
Const strTxtPrompt = "H�zl�"
Const strTxtBasic = "Basit"
Const strTxtAddEmailLink = "E-posta Ekle"
Const strTxtList = "Listele"
Const strTxtCentre = "Ortala"

Const strTxtEnterBoldText = "Koyu Olarak Yaz�lmas�n� �stedi�iniz Yaz�y� Yaz�n"
Const strTxtEnterItalicText = "E�ik Olarak Yaz�lmas�n� �stedi�iniz Yaz�y� Yaz�n"
Const strTxtEnterUnderlineText = "Alt� �izili Olarak Yaz�lmas�n� �stedi�iniz Yaz�y� Yaz�n"
Const strTxtEnterCentredText = "Ortalanmas�n� Yaz�lmas�n� �stedi�iniz Yaz�y� Yaz�n"
Const strTxtEnterHyperlinkText = "Ba�lant� olarak g�r�nt�lenmesini istedi�iniziyaz�n"
Const strTxtEnterHeperlinkURL = "Ba�lant� yap�lmas�n� istedi�iniz URL yi yaz�n"
Const strTxtEnterEmailText = "E-mail ba�lant�s� olarak g�r�nt�lenmesini istedi�inizi yaz�n"
Const strTxtEnterEmailMailto = "Ba�lant� yap�lacak E-mail adresini yaz�n"
Const strTxtEnterImageURL = "Eklemek �stedi�iniz Resmin Web Adresini Yaz�n"
Const strTxtEnterTypeOfList = "Liste Tipi"
Const strTxtEnterEnter = "Gir"
Const strTxtEnterNumOrBlankList = "S�ral� Olmas�n� �stiyorsan�z Bo� B�rak�n"
Const strTxtEnterListError = "HATA! L�tfen Girin"
Const strEnterLeaveBlankForEndList = "Yaz� Listeye Eklendi, Listeyi Sonland�rmak ��in Bo� B�rak�n"
Const strTxtErrorInsertingObject = "Hata �imdiki Konumda Nesne Ekleme."


Const strTxtFontStyle = "Karakter Tipi"
Const strTxtFontTypes = "Font"
Const strTxtFontSizes ="Boyut"
Const strTxtEmoticons = "Duygu Simgeleri"
Const strTxtFontSize = "Font Boyutu"


Const strTxtFontColours ="Font Renkleri"
Const strTxtBlack = "Siyah"
Const strTxtWhite = "Beyaz"
Const strTxtBlue = "Mavi"
Const strTxtRed = "K�rm�z�"
Const strTxtGreen = "Ye�il"
Const strTxtYellow = "Sar�"
Const strTxtOrange = "Turuncu"
Const strTxtBrown = "Kahverengi"
Const strTxtMagenta = "Pembe"
Const strTxtCyan = "A��k Mavi"
Const strTxtLimeGreen = "A��k Ye�il"



Const strTxtCut = "Kes"
Const strTxtCopy = "Kopyala"
Const strTxtPaste = "Yap��t�r"
Const strTxtBold = "Koyu"
Const strTxtItalic = "E�ik"
Const strTxtUnderline = "Alt� �izili"
Const strTxtLeftJustify = "Sola Hizal�"
Const strTxtCentrejustify = "Ortaya Hizal�"
Const strTxtRightJustify = "Sa�a Hizal�"
Const strTxtJustify = "Hizala"
Const strTxtUnorderedList = "D�zensiz Liste"
Const strTxtOutdent = "D��a Do�ru"
Const strTxtIndent = "��e Do�ru"
Const strTxtAddHyperlink = "Ba�lant� Ekle"
Const strTxtAddImage = "Resim Ekle"
Const strTxtJavaScriptEnabled = "Foruma Mesaj Yollayabilmeniz ��in Javascript Etkin Olmal�!"
Const strTxtFontColour = "Renk"
Const strTxtstrTxtOrderedList = "D�zenli Liste"
Const strTxtTextColour = "Yaz� Rengi"
Const strTxtBackgroundColour = "Arkaplan rengi"
Const strTxtUndo = "Geri Al"
Const strTxtRedo = "�leri Al"
Const strTxtstrSpellCheck = "Yaz�m Denetimi"
Const strTxtToggleHTMLView = "Html G�r�nt�s�"
Const strTxtAboutRichTextEditor = "Zengin Metin Edit�r� Hakk�nda"
Const strTxtInsertTable = "Tablo Ekle"
Const strTxtSpecialCharacters = "�zel Karakterler"
Const strTxtPrint = "Yazd�r"
Const strTxtImage = "Resim"
Const strTxtStrikeThrough = "Strike Through"
Const strTxtSubscript = "Subscript"
Const strTxtSuperscript = "Superscript"
Const strTxtHorizontalRule = "Yatay �l��"



Const strTxtIeSpellNotDetected = "ie Kontrol yapamad�. �ndirmek i�in tamam tu�una t�klay�n�z."
Const strTxtSpellBoundNotDetected = "You need \'SpellBound 0.7.0+\' spelling checker installed to use this feature. \nClick OK to go to the \'SpellBound\' download page."



Const strTxtOK = "Tamam"
Const strTxtCancel = "�ptal"


Const strTxtImageUpload = "Resim Y�kle"
Const strTxtFileUpload = "Dosya Y�kle"
Const strTxtUpload = "Y�kle"
Const strTxtPath = "Uzant�"
Const strTxtFileURL = "Dosya URL"

Const strTxtParentDirectory = "Ana Klas�r"

Const strTxtImagesMustBeOfTheType = "Bu t�r resim eklenmeli"
Const strTxtAndHaveMaximumFileSizeOf = "ve boyutu"
Const strTxtImageOfTheWrongFileType = "Dosya t�r� yanl��"
Const strTxtImageFileSizeToLarge = "Bu resmin dosya boyutu kadar olmal�"
Const strTxtMaximumFileSizeMustBe = "Maksimum dosya boyutu �u kadar olmal�"
Const strTxtErrorUploadingImage = "Resim Y�kleme Hatas�!!"
Const strTxtNoImageToUpload = "L�tfen bunu kullan�n \'G�zat...\' sonra y�klemek istedi�iniz resimi se�in."

Const strTxtFile = "Dosya"
Const strTxtFilesMustBeOfTheType = "Dosya uzant�s� �u �ekilde olmal�"
Const strTxtFileOfTheWrongFileType = "Upload edilen dosya t�r� yanl��"
Const strTxtFileSizeToLarge = "Bu kadardan b�y�k dosya"
Const strTxtErrorUploadingFile = "Dosya Y�kleme Hatas�!!"
Const strTxtNoFileToUpload = "L�tfen bunu kullan�n \'G�zat...\' sonra y�klemek istedi�iniz dosyay� se�in."


Const strTxtPleaseWaitWhileFileIsUploaded = "L�tfen dosya servere g�nderilirken bekleyin."
Const strTxtPleaseWaitWhileImageIsUploaded = "L�tfen resim servere g�nderilirken bekleyin."


Const strTxtCloseWindow = "Pencereyi Kapat"


Const strTxtPreview = "�nizleme"
Const strTxtThereIsNothingToPreview = "�nizleme Yap�lamad�"

Const strResetFormConfirm = "Formu Temizlemek �stedi�inizden Eminmisiniz?"
Const strResetWarningFormConfirm = "UYARI: Formdaki t�m bilgiler kaybolacak!!"
Const strResetWarningEditorConfirm = "UYARI: D�zenlenen t�m bilgiler kaybolacak!!"


Const strTxtSubmitForm = "Formu Sun"
Const strTxtResetForm = "Formu Yenile"

Const strTxtDisplayMessage = "Mesaj G�ster"
Const strTxtThereIsNothingToShow = "G�sterilebilecek Mesaj Yok"


Const strTxtTableProperties = "Tablo �zellikleri"

Const strTxtImageProperties = "Resim �zellikleri"

Const strTxtImageURL = "Resim&nbsp;URL"
Const strTxtAlternativeText = "Alternatif Yaz�"
Const strTxtLayout = "Taslak"
Const strTxtAlignment = "Hizalama"
Const strTxtBorder = "�er�eve"
Const strTxtSpacing = "Bo�luklar"
Const strTxtHorizontal = "Yatay"
Const strTxtVertical = "Dikey"

Const strTxtRows = "S�ra"
Const strTxtColumns = "S�tunlar"
Const strTxtWidth = "Geni�lik"
Const strTxtpixels = "pixel"
Const strTxtCellPad = "H�cre �er�evesi"
Const strTxtCellSpace = "H�cre bo�lu�u"

Const strTxtHeight = "Y�kseklik"


Const strTxtSelectTextToTurnIntoHyperlink = "L�tfen ba�lant�ya �evrilecek birka� yaz� se�in."

Const strTxtYourBrowserSettingsDoNotPermit = "Taray�c�n�z edit�r�n isteklerine izin vermiyor"
Const strTxtPleaseUseKeybordsShortcut = "operations. \nL�tfen klavye k�sayollar�n� kullan�n "
Const strTxtWindowsUsers = "Windows kullan�c�lar�: "
Const strTxtMacUsers = "Mac kullan�c�lar�: "


Const strTxtHyperlinkProperties = "Ba�lant� �zellikleri"
Const strTxtNoPreviewAvailableForLink = "�nizleme mevcut de�il"
Const strTxtAddress = "Adres"
Const strTxtLinkType = "Link T�r�"
Const strTxtTitle = "Ba�l�k"
Const strTxtWindow = "Pencere"
Const strTxtEmail = "Eposta"
Const strTxtSubject = "Konu"
Const strTxtPleaseWaitWhilePreviewLoaded = "L�tfen bekleyin, �nizleme y�kleniyor...."
Const strTxtErrorLoadingPreview = "�nizleme Y�kleme Hatas�.\nL�tfen uzant�y� ve ismi kontrol edin."


Const strTxAttachFileProperties = "Dosya �zellikleri"

Const strTxtNewBlankDoc = "Yeni Bo� D�k�man"
Const strTxtOpen = "A�"
Const strTxtSave = "Kaydet"




Const strTxtPasteFromWord = "Word dan yap��t�r"
Const strTxtPasteFromWordDialog = "Bu form Word'en gelen yaz�lar� temzilemek i�in kullan�l�r. L�tfen a�a��daki kutucu�a klavyenizi kullanarak (Windows kullan�c�lar�: Ctrl + 'v', MAC kullan�c�lar: Apple + 'v') kopyalad���n�z metni yap��t�r�n ve 'Tamam' butonuna bas�n."

Const strTxtFileAlreadyExistsRenamedAs = "Ayn� isme sahip iki dosya var veya girmi� oldu�unuz dosyan�n isminde bir problem var.\nDosya �u �ekilde kaydedildi:"
Const strTxtTheFile = "Dosya:"
Const strTxtHasBeenSaved = "kaydedildi"

%>