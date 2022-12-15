<% @ Language = VBScript %>
<%
Option Explicit




Response.Buffer = True
Response.Expires = 0
Response.Clear

Class FileUpload
	Private UploadRequest, oProps, iFrmCt
	Private iKnownFileCount, iKnownFormCount	
	Private oOutFiles

	Private Sub Class_Initialize
		iFrmCt = 0
		Set oProps = New FO_Properties
		Set UploadRequest = Server.CreateObject("Scripting.Dictionary")
		iKnownFileCount = 0
		iKnownFormCount = 0
		set oOutFiles = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate
		set oOutFiles = Nothing
		Set UploadRequest = Nothing
		Set oProps = Nothing
	End Sub

	Public Property Get Version()
		Version = "1.0"
	End Property

	Public Function GetUploadSettings()
		Set GetUploadSettings = oProps
	End Function

	Public Property Get FormCount
		FormCount = iKnownFormCount
	End Property

	Public Property Get FileCount
		FileCount = iKnownFileCount
	End Property

	Public Property Get TotalFormCount
		TotalFormCount = iFrmCt
	End Property

	Private Function GetFormEncType()
		Dim sContType, hCutOff

		sContType = Request.ServerVariables("CONTENT_TYPE")
		hCutOff = instr(sContType, ";")
		if hCutOff > 0 Then
			sContType = UCase(Trim(Left(sContType, hCutOff - 1)))
		else
			sContType = UCase(Trim(sContType))
		End If
		GetFormEncType = sContType
	End Function

	Public Default Sub ProcessUpload
		Dim RequestBin, oProcess, iTotBytes, key, arr, iKnownProps, oFile
		Dim fofilecheck, sEncType, sReqMeth

		iTotBytes = Request.TotalBytes
		if iTotBytes = 0 Then
			iFrmCt = 0
			exit sub
		End If
		RequestBin = Request.BinaryRead(iTotBytes)

		sReqMeth = Request.ServerVariables("REQUEST_METHOD")
		select case UCase(sReqMeth)
			case "POST"
				sEncType = GetFormEncType
				select case sEncType
					case "MULTIPART/FORM-DATA"

						Set oProcess = New FO_Processor
						oProcess.BuildUploadRequest  RequestBin, UploadRequest
						Set oProcess = Nothing

					case "APPLICATION/X-WWW-FORM-URLENCODED"

				
						Set oProcess = New FO_Processor
						oProcess.BuildUploadRequest_ASCII oProcess.getString(RequestBin), UploadRequest
						Set oProcess = Nothing

					case else

				
				end select

			case "GET"
			
			case else
				
		end select

		arr = uploadrequest.keys

		if not isarray(arr) Then
			iFrmCt = 0
			exit sub
		End If

		iFrmCt = ubound(arr)
		for each key in arr
			if isobject(uploadrequest.item(key)) Then
				iKnownProps = ubound(uploadrequest.item(key).keys) + 1
				if iKnownProps = 4 Then
					iKnownFileCount = iKnownFileCount + 1
					set fofilecheck = new FO_FileChecker
					fofilecheck.SetCurrentProperties oProps
					fofilecheck.FileInput_NamePath = uploadrequest.item(key).item("FileName")
					fofilecheck.FileInput_ContentType = uploadrequest.item(key).item("ContentType")
					fofilecheck.FileInput_BinaryText = uploadrequest.item(key).item("Value")
					fofilecheck.FileInput_FormInputName = uploadrequest.item(key).item("InputName")
					set oFile = fofilecheck.ValidateVerifyReturnFile()
					set fofilecheck = nothing

					oOutFiles.add iKnownFileCount, oFile
					set oFile = nothing
					uploadrequest.remove key
				elseif iKnownProps = 2 Then
					iKnownFormCount = iKnownFormCount + 1
				else
					End If
			End If
		next
	End Sub

	Public Function File(ByVal blobName)
		Dim blobs, blob, subdict, tmpName
		blobs = oOutFiles.Keys
		For Each blob In blobs
			Set subdict = oOutFiles.Item(blob)
			tmpName = subdict.frmInputName
			If UCase(Trim(tmpName)) = UCase(Trim(blobName)) Then
				blobName = blob
				Exit For
			End If
		Next
		if isobject(oOutFiles.Item(blobName)) Then
			Set File = oOutFiles.Item(blobName)
		else
			Set File = Nothing
		End If
	End Function

	Public Function Form(ByVal inputName)
		if isobject(UploadRequest.Item(inputName)) Then
			Form = UploadRequest.Item(inputName).Item("Value")
		else
			Form = ""
		End If
	End Function

	Public Function FormLen(ByVal inputName)
		if isobject(UploadRequest.Item(inputName)) Then
			FormLen = Len(UploadRequest.Item(inputName).Item("Value"))
		else
			FormLen = 0
		End If
	End Function

	Public Function FormEx(ByVal inputName, ByVal vDefaultValue)
		dim vTmp

		if isobject(UploadRequest.Item(inputName)) Then
			vTmp = UploadRequest.Item(inputName).Item("Value")
			if len(trim(CStr(vTmp))) = 0 Then
				FormEx = vDefaultValue
				Exit Function
			End If

			FormEx = vTmp
			Exit Function
		End If

		FormEx = vDefaultValue
	End Function

	Public Function Inputs()
		if isobject(UploadRequest) Then
			Inputs = UploadRequest.keys
		else
			Inputs = ""
		End If
	End Function

	Public Sub ShowUploadForm(ByVal sSubmitPage)
		Dim tmp, item

		With Response

			.Write("Max. Dosya boyutu: <CODE>~ ")
			.Write(Round( oProps.MaximumFileSize / 1024, 1 ) & " Kb.</CODE> ")


			.Write("<FORM ENCTYPE=""multipart/form-data"" ACTION=""")
			.Write(sSubmitPage & """ METHOD=""POST"">" & vbCrLf)

			.Write("Lütfen bir dosya seçin")
			if oProps.UploadDisabled Then
				.Write("Bilgisayarýnýzdan dosya yüklemeniz imkansýz:<BR>" & vbCrLf)
				.Write("<INPUT TYPE=FILE NAME=""blob"" DISABLED><BR><BR>" & vbCrLf)
			Else
				.Write(":")
				.Write("")

				.Write("<BR>" & vbCrLf)
				.Write("<INPUT TYPE=""FILE"" NAME=""blob"" src=""xx"" class=""files"" style=""width: 200px;border:1px solid #CCC;margin: 5px 0 0 0;""><BR><BR>" & vbCrLf)
			End If

			
			.Write("<INPUT NAME=""myName"" type=""Hidden"" >" & vbCrLf)
			.Write("<INPUT TYPE=""SUBMIT"" VALUE=""Yükle"">" & vbCrLf)
			.Write("</FORM>" & vbCrLf)
		End With
	End Sub
End Class



Class FO_FileChecker
	Private oProps, sFileName, hFileBinLen, sFileBin, sFileContentType, sFileFormInputName

	Private Sub Class_Initialize()
		sFileName = ""
		hFileBinLen = 0
		sFileBin = ""
		sFileContentType = ""
	End Sub

	Public Sub SetCurrentProperties(byref oPropertybag)
		Set oProps = oPropertybag
	End Sub

	Public Property Let FileInput_FormInputName(ByVal fname)
		sFileFormInputName = fname
	End Property

	Public Property Let FileInput_NamePath(ByVal fname)
		Dim realfilename

		realfilename = Right(fname, Len(fname) - InstrRev(fname,"\"))

		sFileName = trim(realfilename)
	End Property

	Public Property Let FileInput_ContentType(ByVal conttype)
		sFileContentType = conttype
	End Property

	Public Property Let FileInput_BinaryText(ByVal binstring)
		Dim  binlen

		binlen = lenb(binstring)
		hFileBinLen = binlen
		sFileBin = binstring
	End Property

	Public Function ValidateVerifyReturnFile()	
		if IllegalCharsFound Then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "dosya adýnda geçersiz karakter bulunamaz", "", "", "", sFileFormInputName)
			Exit Function
		End If

		if FileNameBadOrExists Then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "bir dosya seçmediniz ya da seçtiðiniz dosya yolu yanlýþ; bir diðer olasýlýk seçtiðiniz dosya zaten yüklü", "", "", "", sFileFormInputName)
			Exit Function
		End If

		If FileExtensionIsBad Then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "bu dosya türü desteklenmemektedir", "", "", "", sFileFormInputName)
			Exit Function
		End If

		If FileSizeIsBad Then
			Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "dosya boyutu uygun boyutta deðil. lütfen max. ve min. boyutlar arasýnda bir dosya yükleyiniz.", "", "", "", sFileFormInputName)
			Exit Function
		End If

		Set ValidateVerifyReturnFile = FillFOFileObj(false, "", "", "", sFileContentType, sFileName, sFileBin, sFileFormInputName)
	End Function

	Private Function FillFOFileObj(byval success, byval abspath, byval virpath, byval stderr, byval contenttype, byval fname, byval binarytext, byval forminputname)
		Dim oFile
		set oFile = New FO_File
		oFile.SetCurrentProperties oProps
		oFile.bSuccess = success
		oFile.sAbsPath = abspath
		oFile.sVirPath = virpath
		oFile.sStdErr = stderr
		oFile.sCType = contenttype
		oFile.sFileName = fname
		oFile.binValue = binarytext
		oFile.frmInputName = forminputname
		set FillFOFileObj = oFile
	End Function	

	Public Function IllegalCharsFound()
		Dim re

		set re = new regexp
		re.pattern = "\\\/\:\*\?\""\<\>\|" ' burada hackerlara engel koyuyoruz
		re.global = true
		re.ignorecase = true
		if re.test(sFileName) Then
			IllegalCharsFound = true
		else
			IllegalCharsFound = false
		End If
		set re = nothing
	End Function

	Public Function FileNameBadOrExists()
		Dim absuploaddirectory, oFSO

		if len(trim(sFileName)) = 0 Then
			FileNameBadOrExists = true
			Exit Function
		End If
		
		if oProps.AllowOverWrite Then
			FileNameBadOrExists = false
			Exit Function
		End If

		absuploaddirectory = oProps.uploaddirectory & "\" & trim(sFileName)

		set oFSO = server.createobject("Scripting.FileSystemObject")
		if oFSO.FileExists(absuploaddirectory) Then
			FileNameBadOrExists = true
		else
			FileNameBadOrExists = false
		End If
		Set oFSO = Nothing
	End Function

	Public Function FileExtensionIsBad()
		Dim sFileExtension, bFileExtensionIsValid, sFileExt

		if len(trim(sFileName)) = 0 Then
			FileExtensionIsBad = true
			Exit Function
		End If

		sFileExtension = right(sFileName, len(sFileName) - instrrev(sFileName, "."))
		bFileExtensionIsValid = false	
		for each sFileExt in oProps.extensions
			if ucase(sFileExt) = ucase(sFileExtension) Then
				bFileExtensionIsValid = True
				exit for
			End If
		next
		FileExtensionIsBad = False
	End Function

	Public Function FileSizeIsBad()
		if hFileBinLen > oProps.MaximumFileSize Then
			FileSizeIsBad = True
			Exit Function
		End If

		if hFileBinLen < oProps.MininumFileSize Then
			FileSizeIsBad = True
			Exit Function
		End If

		FileSizeIsBad = False
	End Function
End Class



Class FO_Processor
	Private Function getByteString(byval StringStr)
		dim char, i

		For i = 1 to Len(StringStr)
			char = Mid(StringStr, i, 1)
			getByteString = getByteString & chrB(AscB(char))
		Next
	End Function

	Public Function getString(byval StringBin)
		dim intCount

		getString =""
		For intCount = 1 to LenB(StringBin)
			getString = getString & chr(AscB(MidB(StringBin, intCount, 1))) 
		Next
	End Function

	Public Sub BuildUploadRequest_ASCII(ByVal sPostStr, ByRef UploadRequest) 
		dim i, j, blast, sName, vValue
		dim tmphash

		blast = false
		i = -1
		do while i <> 0
			if i = -1 Then
				i = 1
			else
				i = i + 1
			End If
			j = instr(i, sPostStr, "=") + 1
			sName = mid(sPostStr, i, j-i-1)
			i = instr(j, sPostStr, "&")
			if i = 0 Then 
				vValue = mid(sPostStr, j)
			else
				vValue = mid(sPostStr, j, i - j)
			End If

			Dim uploadcontrol
			set uploadcontrol = createobject("Scripting.Dictionary")
			uploadcontrol.add "Value", vValue

			if not uploadrequest.exists(sName) Then
				uploadrequest.add sName, uploadcontrol
			else
				set tmphash = uploadrequest(sName)
				tmphash("Value") = tmphash("Value") & ", " & vValue
				set uploadrequest(sName) = tmphash
			End If
		loop
	End Sub



	Public Sub BuildUploadRequest(byref RequestBin, byref UploadRequest)
		dim PosBeg, PosEnd, boundary, boundaryPos, Pos, Name, PosFile
		dim PosBound, FileName, ContentType, Value, sEncType, sReqMeth
		dim tmphash, isfile

		if lenb(RequestBin) = 0 Then 
			exit sub
		End If

		PosBeg = 1
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))

		if posend = 0 Then
			BuildUploadRequest_ASCII getString(requestbin), UploadRequest
			Exit Sub
		End If

		boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		boundaryPos = InstrB(1,RequestBin,boundary)
		Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
			Dim UploadControl
			Set UploadControl = Server.CreateObject("Scripting.Dictionary")
			Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
			Pos = InstrB(Pos,RequestBin,getByteString("name="))
			PosBeg = Pos+6
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
			PosBound = InstrB(PosEnd,RequestBin,boundary)

			isfile = false

			If  PosFile<>0 AND (PosFile<PosBound) Then
				PosBeg = PosFile + 10
				PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
				FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "FileName", FileName
				Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
				PosBeg = Pos+14
				PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
				ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
				UploadControl.Add "ContentType",ContentType
				PosBeg = PosEnd+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)

				isfile = true
			Else
				Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
				PosBeg = Pos+4
				PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
				Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))

				isfile = false
			End If
			UploadControl.Add "Value" , Value
			UploadControl.Add "InputName", Name
			if not uploadrequest.exists(name) Then 
				UploadRequest.Add name, UploadControl	
			else
				if not isfile Then
					set tmphash = uploadrequest(name)
					tmphash("Value") = tmphash("Value") & ", " & Value
					set uploadrequest(name) = tmphash
				End If
			End If

			BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
		Loop
	End Sub
End Class



Class FO_File
	Public bSuccess
	Public sAbsPath
	Public sVirPath
	Public sStdErr
	Public sCType
	Public frmInputName
	Public binValue
	Private hBtCt, sURiPath, sFiExt
	private sfinme

	Private oProps

	Public property let sFileName(byval filenameinput)
		sFiExt = right(filenameinput, len(filenameinput) - instrrev(filenameinput, "."))
		sfinme = filenameinput
	end property

	public property get sFileName()
		sFileName = sfinme
	end property

	Private Sub Class_Initialize()
		bSuccess = false
		sAbsPath = ""
		sVirPath = ""
		sStdErr = ""
		hBtCt = 0
		sCType = ""
		sFileName = ""
		binValue = ""
		sURiPath = ""
	End Sub

	Public Sub SetCurrentProperties(byref oPropertybag)
		Set oProps = oPropertybag
	End Sub

	Public Sub SaveAsRecord(byref oField)
		sAbsPath = ""
		sVirPath = ""
		sURiPath = ""
		bSuccess = false

		If LenB(binValue) = 0 Then 
			Exit Sub
		End If

		if oProps.UploadDisabled Then
			sStdErr = "Uploading disabled by administrator"
			Exit Sub
		End If
		
		If IsObject(oField) Then
			On Error Resume Next
			oField.AppendChunk binValue
			if Err Then
				sStdErr = Err.Description
				bBtCt = 0
				bSuccess = false
				Exit Sub
			End If
			On Error GoTo 0

			hBtCt = lenb(binValue)
			bSuccess = true
		End If
	End Sub

	Public Sub SaveAsFile()
		If sStdErr <> "" Then
			exit sub
		End If
		WriteUploadFile oProps.uploaddirectory & "\" & sFileName, binValue
	End Sub

	Public Function SaveAsBinaryString()
		If LenB(binValue) = 0 Then 
			bBtCt = 0
			bSuccess = false
			Exit Function
		End If

		if oProps.UploadDisabled Then
			bBtCt = 0
			bSuccess = false
			sStdErr = "Yükleme Sayfa Yönetimi Tarafýndan Engellendi"
			Exit Function
		End If

		SaveAsBinaryString = binValue
		hBtCt = lenb(binValue)
		bSuccess = true
	End Function

	Public Function SaveAsString()
		Dim outstr, i

		If LenB(binValue) = 0 Then 
			bBtCt = 0
			bSuccess = false
			Exit Function
		End If

		if oProps.UploadDisabled Then
			bBtCt = 0
			bSuccess = false
			sStdErr = "Yükleme Sayfa Yönetimi Tarafýndan Engellendi"
			Exit Function
		End If

		outstr = ""
		For i = 1 to LenB( binValue )
			outstr = outstr & chr( AscB( MidB( binValue, i, 1) ) )
		Next
		SaveAsString = outstr
		hBtCt = lenb(binValue)
		bSuccess = true
	End Function

	Public Function SaveAsBase64EncodedStr()
		Dim outstr, oEnc

		If LenB(binValue) = 0 Then 
			bBtCt = 0
			bSuccess = false
			Exit Function
		End If

		if oProps.UploadDisabled Then
			bBtCt = 0
			bSuccess = false
			sStdErr = "Yükleme Sayfa Yönetimi Tarafýndan Engellendi"
			Exit Function
		End If
		Set oEnc = New Base64Encoder
		outstr = oEnc.EncodeStr(binValue)
		Set oEnc = Nothing
		SaveAsBase64EncodedStr = outstr
		hBtCt = lenb(binValue)
		bSuccess = true
	End Function

	Private Sub WriteUploadFile(byVal NAME, byVal CONTENTS)
		dim ScriptObject, i, NewFile

		on error resume next

		if oProps.UploadDisabled Then
			err.raise "31234", "FO Obj", "Yükleme Sayfa Yönetimi Tarafýndan Engellendi"
		else
			Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
			Set NewFile = ScriptObject.CreateTextFile( NAME )
			For i = 1 to LenB( CONTENTS )
			NewFile.Write chr( AscB( MidB( CONTENTS, i, 1) ) )
			Next
			NewFile.Close
			Set NewFile = Nothing
			Set ScriptObject = Nothing
		End If
		if err.number <> 0 Then
			sStdErr = Err.Description
			bSuccess = false
		else
			sAbsPath = NAME
			sVirPath = UnMappath(NAME)
			hBtCt = lenb(CONTENTS)
			sURiPath = "http://" & Request.ServerVariables("HTTP_HOST") & sVirPath
			bSuccess = true
		End If
		on error goto 0
	End Sub

	Private Function UnMappath(byVal pathname)
		dim tmp, strRoot

		strRoot = Server.Mappath("/")
		tmp = replace( lcase( pathname ), lcase( strRoot ), "" )
		tmp = replace( tmp, "\", "/" )
		UnMappath = tmp
	End Function

	Public Property Get ContentType()
		ContentType = sCType
	End Property

	Public Property Let FileName(byval newfilename)
		Dim oFileChk
		set oFileChk = New FO_FileChecker
		oFileChk.SetCurrentProperties oProps
		oFileChk.FileInput_NamePath = newfilename
		if oFileChk.IllegalCharsFound Then
			sStdErr = "Dosya içerisinde geçersiz karakterler bulundu"
			bSuccess = false
			set oFileChk = Nothing
			Exit Property
		End If
		if oFileChk.FileNameBadOrExists Then
			sStdErr = "Dosya ismi geçersiz ya da bu dosyadan zaten mevcut ve üstüne yazma engellenmiþ"
			bSuccess = false
			set oFileChk = Nothing
			Exit Property
		End If
		if oFileChk.FileExtensionIsBad Then
			sStdErr = "bu dosya türü desteklenmemektedir"
			bSuccess = false
			set oFileChk = Nothing
			Exit Property
		End If
		Set oFileChk = Nothing

		sStdErr = ""
		sFileName = newfilename
	End Property

	Public Property Get FileExtension()
		FileExtension = sFiExt
	End Property

	Public Property Get FileNameWithoutExtension()
	FileNameWithoutExtension = StripFileExtensionFromFileName(sFileName)
	End Property

	Public Function StripFileExtensionFromFileName(ByVal filenametostrip)
		Dim hExtensionStart, tmpfilenametoalter

		tmpfilenametoalter = filenametostrip
		hExtensionStart = -1
		do while not hExtensionStart = 0
			hExtensionStart = instrrev(tmpfilenametoalter, ".")
			if hExtensionStart > 0 Then
				tmpfilenametoalter = left(tmpfilenametoalter, hExtensionStart - 1)
			End If
		loop
		StripFileExtensionFromFileName = tmpfilenametoalter
	End Function

	Public Function JoinFileExtensionToFileName(ByVal filenametojoin, byval fileextensiontojoin)
		Dim strippedfilename

		strippedfilename = StripFileExtensionFromFileName(filenametojoin)
		JoinFileExtensionToFileName = strippedfilename & "." & fileextensiontojoin
	End Function

	Public Function GetFileNameFromFilePath(ByVal filewithpath)
		dim fileend

		fileend = instrrev(filewithpath, "\")
		GetFileNameFromFilePath = right(filewithpath, len(filewithpath) - fileend)
	End Function

	Public Property Get FileName()
		FileName = sFileName
	End Property

	Public Property Get UploadSuccessful()
		UploadSuccessful = bSuccess
	End Property

	Public Property Get AbsolutePath()
		AbsolutePath = sAbsPath
	End Property

	Public Property Get URLPath()
		URLPath = sURiPath
	End Property

	Public Property Get VirtualPath()
		VirtualPath = sVirPath
	End Property

	Public Property Get ErrorMessage()
		ErrorMessage = sStdErr
	End Property

	Public Property Get ByteCount()
		ByteCount = hBtCt
	End Property
End Class



Class FO_Properties
	Private sErrHead		
	Private sErrMsg			
	Private arrExt			
	Private strUploadDir		
	Private boolAllowOverwrite	
	Private lngUploadSize		
	Private bMin			
	Private bByPass			

	Private Sub Class_Initialize()
		sErrHead = "Yanlýþ Kurulum Hatasý"
		sErrMsg = ""
		arrExt = Array("tar", "gz", "zip", "tgz") ' dosya uzantýlarý DÝKKAT
		strUploadDir = Server.Mappath("/")
		boolAllowOverwrite = false
		lngUploadSize = 100000  
		bMin = 1024
		bByPass = false
	End Sub

	Public Sub ResetAll()
		Class_Initialize
	End Sub

	Public Property LET Extensions(byVal arrayInput)
		dim item, bErr

		bErr = false
		if isarray(arrayInput) Then
			for each item in arrayInput
				if instr(item, ".") <> 0 Then
					bErr = true
					exit for
				End If
			next
			if not bErr Then
				arrExt = arrayInput
				Exit Property
			else
				arrayInput = ""
			End If
		End If

		sErrMsg = "ASP dosyasýnda bulunan uzantýlara nokta koymamalýsýnýz(.)."
		if arrayInput = "*" Then
			Err.Raise 21340, sErrHead, sErrMsg & _
				" Desteklenmiyor."
		else
			Err.Raise 21341, sErrHead, sErrMsg
		End If
	End Property

	Public Property LET UploadDirectory(byVal strInput)
		Dim oFSO, bDoesntExist

		bDoesntExist = false

		if instr(strInput, "/") <> 0 Then
			strInput = ""
			Err.Raise 21342, sErrHead, _
				"Veri yolu tam olarak girilmeli."
			exit property
		End If

		Set oFSO = CreateObject("Scripting.FileSystemObject")
		if not oFSO.FolderExists(strInput) Then bDoesntExist = true
		set oFSO = Nothing
		if bDoesntExist Then
			Err.Raise 21343, sErrHead, "HATA - """ & _
				strInput & """ Bu dosya serverda bulunmamaktadýr."
			Exit Property
		End If

		strUploadDir = strInput
	End Property

	Public Property LET AllowOverWrite(byVal boolInput)
		on error resume next
		boolInput = cbool(boolInput)
		on error goto 0
		boolAllowOverwrite = boolInput
	End Property

	Public Property LET MaximumFileSize(byVal lngInput)
		if isnumeric(lngInput) Then
			on error resume next
			lngInput = CLng( lngInput )
			on error goto 0

			lngUploadSize = lngInput
			exit property
		End If

		Err.Raise 21344, sErrHead, "Maksimum dosya boyutu rakamlardan oluþmalýdýr."
	End Property

	Public Property LET MininumFileSize(byVal lngInput)
		if isnumeric(lngInput) Then
			on error resume next
			lngInput = CLng( lngInput )
			on error goto 0

			bMin = lngInput
			exit property
		End If

		Err.Raise 21345, sErrHead, "Minimum dosya boyutu rakamlardan oluþmalýdýr."
	End Property

	Public Property LET UploadDisabled(byval boolInput)
		on error resume next
		boolInput = cbool(boolInput)
		on error goto 0
		bByPass = boolInput
	End Property

	Public Property GET UploadDisabled()
		UploadDisabled = bByPass
	End Property

	Public Property GET MininumFileSize()
		MininumFileSize = bMin
	End Property

	Public Property GET Extensions()
		Extensions = arrExt
	End Property

	Public Property GET UploadDirectory()
		UploadDirectory = strUploadDir
	End Property

	Public Property GET AllowOverWrite()
		AllowOverWrite = boolAllowOverwrite
	End Property

	Public Property GET MaximumFileSize()
		MaximumFileSize = lngUploadSize
	End Property
End Class

Class Base64Encoder
	Private Base64Chars

	Private Sub Class_Initialize()
		Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
				"abcdefghijklmnopqrstuvwxyz" & _
				"0123456789" & _
				"+/"
	End Sub

	Public Function EncodeStr(byVal strIn)
		Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 3
			c1 = Asc(Mid(strIn, n, 1))
			c2 = Asc(Mid(strIn, n + 1, 1) + Chr(0))
			c3 = Asc(Mid(strIn, n + 2, 1) + Chr(0))
			w1 = Int(c1 / 4) : w2 = (c1 And 3) * 16 + Int(c2 / 16)
			If Len(strIn) >= n + 1 Then 
				w3 = (c2 And 15) * 4 + Int(c3 / 64) 
			Else 
				w3 = -1
			End If
			If Len(strIn) >= n + 2 Then 
				w4 = c3 And 63 
			Else 
				w4 = -1
			End If
			strOut = strOut + mimeencode(w1) + mimeencode(w2) + _
					  mimeencode(w3) + mimeencode(w4)
		Next
		EncodeStr = strOut
	End Function

	Private Function mimedecode(byVal strIn)
		If Len(strIn) = 0 Then 
			mimedecode = -1 : Exit Function
		Else
			mimedecode = InStr(Base64Chars, strIn) - 1
		End If
	End Function

	Public Function DecodeStr(byVal strIn)
		Dim w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 4
			w1 = mimedecode(Mid(strIn, n, 1))
			w2 = mimedecode(Mid(strIn, n + 1, 1))
			w3 = mimedecode(Mid(strIn, n + 2, 1))
			w4 = mimedecode(Mid(strIn, n + 3, 1))
			If w2 >= 0 Then _
				strOut = strOut + _
					Chr(((w1 * 4 + Int(w2 / 16)) And 255))
			If w3 >= 0 Then _
				strOut = strOut + _
					Chr(((w2 * 16 + Int(w3 / 4)) And 255))
			If w4 >= 0 Then _
				strOut = strOut + _
					Chr(((w3 * 64 + w4) And 255))
		Next
		DecodeStr = strOut
	End Function


	Private Function mimeencode(byVal intIn)
		If intIn >= 0 Then 
			mimeencode = Mid(Base64Chars, intIn + 1, 1) 
		Else 
			mimeencode = ""
		End If
	End Function
End Class

If Session("durum")="esp" Then
%><style type="text/css">
body {
	font-family: arial;
	font-size: 12px;
	color: #333;
}
input {
	width: 200px;
	border:1px solid #CCC;
	margin: 5px 0 0 0;
}
button{	width: 200px;
	border:1px solid #CCC;
	margin: 5px 0 0 0;
}
.files {
	width: 108px;
	border:1px solid #CCC;
	margin: 5px 0 0 0;
}
select {
	border:1px solid #CCC;
	margin: 0;
}
</style>
<%Dim oFO, oProps, oFile, i, item, oMyName

Set oFO = New FileUpload

Set oProps = oFO.GetUploadSettings
with oProps
	.Extensions = Array("jpg", "gif", "psd") ' kabul edilen uzantýlar
	.UploadDirectory = Server.Mappath("../Uploads") ' dosyanýn yükleneceði yer
	.AllowOverWrite = true
	.MaximumFileSize = 524288000  ' yüklenmesini istediðiniz maksimum dosya büyüklüðü
	.MininumFileSize = 1 ' burada minimum dikkat ederseniz 1k 1000 diye yazýlýyor 
	.UploadDisabled = false
End with
set oProps = nothing
oFO.ProcessUpload
if oFO.TotalFormCount > 0 Then
	if oFO.FileCount > 0 Then
		for i = 1 to oFO.FileCount
			set oFile = oFO.File(i)
			
			if oFile.ErrorMessage <> "" Then
				Response.Write "&gt; HATA: " & _
					oFile.ErrorMessage & "<BR>"
			else

				oFile.SaveAsFile
				if oFile.UploadSuccessful Then
					Response.Write "&gt; Basariyla Yüklendi<BR>"

					Response.Write(" - Dosyanin su an bulunduðu URL:<font color=""red""> " & _
						oFile.URLPath & "</font><BR>")

					Response.Write(" - Dosya tipi: " & oFile.ContentType & "<BR>")

					Response.Write(" - Dosya ismi: " & oFile.FileName & "<BR>")

					
					Response.Write(" - Dosya boyutu: " & _
						round(formatnumber(oFile.ByteCount, 0)/1024,2) & " KByte<BR>")
				else
					Response.Write "&gt; Dosyayý yüklerken hata oluþtu: " & _
						oFile.ErrorMessage & "<BR>"
				End If
			End If
			set oFile = Nothing
		next
	else
		Response.Write "&gt; Daha önceden bu dosya ile ayni boyutta dosya yüklenmis. Bu durumda ayni dosyayi yüklüyor olabilirsiniz. Eger farkli bir dosya olduguna eminseniz; Dosya boyutunu büyültmek için küçük bir text dosyasini doldurarak zip'li dosyaya ekleyiniz."
	End If


	Response.Write "<BR><BR><A HREF=""" & _
		Request.ServerVariables("SCRIPT_NAME") & """>Tekrar Yükle</A>"
else

	oFO.ShowUploadForm Request.ServerVariables("SCRIPT_NAME")
End If

set oFO = Nothing
Else
Response.Redirect("default.asp")
End If
%>