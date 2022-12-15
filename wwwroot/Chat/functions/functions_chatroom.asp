<%

Private Function postMessage(ByVal strUsername, ByVal strMessage, ByVal strFont, ByVal strColor, ByVal strFormat)

	'Dimesion variables
	Dim strTmpMessage
	Dim saryMessages
	Dim saryTempMessages
	Dim intArrayPass
	Dim intLastMessageID
	Dim lngMessageIndex
	Dim dtmLastMessageTime

	'Initialise the last message date and time
	dtmLastMessageTime = CDbl(NOW())

	'Update the users last post time
	saryWebChatUsers(10, lngLoggedInUserID) = CDbl(NOW())

	'Update the chatroom
	Call updateChatroom()

	'Lock the application so that no other user can try and update the application level variable at the same time
	Application.Lock
	
	'Read in the array from the application variable
	saryMessages = Application("ChatsarryAppChatMessages")

	'Check if there are messages
	If NOT isArray(saryMessages) Then
		
		'ReDimesion the array
		ReDim saryMessages(4, 0)

		'Read in the messages to the application variable
		Application("ChatsarryAppChatMessages") = saryMessages

	End If

	'Remove HTML if any
	strMessage = removeTags(strMessage)

	'format the message
	strMessage = formatMessage(strMessage)

	'Format the style
	If (NOT isNothing(strFont) OR NOT isNothing(strColor) OR NOT isNothing(strFormat)) AND Mid(strMessage, 1, 1) <> "/"  Then

		'Start the style tag
		strTmpMessage = "<span style="""

		'Add the font
		If NOT isNothing(strFont) AND strFont <> strTxtSelectFont Then
		Select Case strFont
		Case "Arial","Book Antiqua","Bookman Old Style","Broadway","Century Gothic","Comic Sans MS","Courier","Garamond","Gill Sans MT","Haettenschweiler",	"Helvetica","Impact","Lucida Bright","Lucida Console","Lucida Sans","Tahoma","Times New Roman","Verdana"
		strTmpMessage = strTmpMessage & "font-family: " & strFont & ";"
		End Select
		End If
		'Add the color
		If NOT isNothing(strColor) Then
		If strColor="red" or strColor="blue" or strColor="green" or strColor="black" or strColor="white" or strColor="violet" or strColor="orange" or strColor="brown" Then
		strTmpMessage = strTmpMessage & "color: " & strColor & ";"
		End If
		End If
		'Add the format
		If strFormat = "i" Then strTmpMessage = strTmpMessage & "font-style: italic;"
		If strFormat = "b" Then strTmpMessage = strTmpMessage & "font-weight: bold;"
		If strFormat = "u" Then strTmpMessage = strTmpMessage & "text-decoration: underline;"

		'End the style tag
		strTmpMessage = strTmpMessage & """>" & strMessage & "</span>"

		'Save the changes
		strMessage = strTmpMessage

	End If

	If NOT isNothing(strMessage) AND strMessage = "/clear" OR NOT isNothing(strMessage) AND Mid(strMessage, 1, 1) <> "/" Then
		
		'Array dimension lookup table
		'0 = Username
		'1 = Message ID
		'2 = Date
		'3 = Message
		'4 = Chatroom ID

		'Read in the size of the array
		lngMessageIndex = UBound(saryMessages, 2)

		'Read in the last message ID
		If lngMessageIndex = 0 Then
			intLastMessageID = 1
		Else
			intLastMessageID = CLng(saryMessages(1, lngMessageIndex)) + 1
		End If

		'Increase by 1
		lngMessageIndex = lngMessageIndex + 1

		'ReDimesion the array
		ReDim Preserve saryMessages(4, lngMessageIndex)

		'Add the message to the array
		saryMessages(0, lngMessageIndex) = strUsername
		saryMessages(1, lngMessageIndex) = intLastMessageID
		saryMessages(2, lngMessageIndex) = dtmLastMessageTime
		saryMessages(3, lngMessageIndex) = strMessage
		saryMessages(4, lngMessageIndex) = intLoggedInUserroom

		'Read in the messages to the application variable
		Application("ChatsarryAppChatMessages") = saryMessages

		'Trim the down to 10 messages
		If UBound(saryMessages, 2) => 10 Then
		
			'ReDimesion the array
			ReDim saryTempMessages(4, 0)
		
			'Loop through the array
			For intArrayPass = UBound(saryMessages, 2) - 10 To UBound(saryMessages, 2)
		
				'ReDimesion the array
				ReDim Preserve saryTempMessages(4, UBound(saryTempMessages, 2) + 1)
		
				'Swap the array positions
				saryTempMessages(0, UBound(saryTempMessages, 2)) = saryMessages(0, intArrayPass)
				saryTempMessages(1, UBound(saryTempMessages, 2)) = saryMessages(1, intArrayPass)
				saryTempMessages(2, UBound(saryTempMessages, 2)) = saryMessages(2, intArrayPass)
				saryTempMessages(3, UBound(saryTempMessages, 2)) = saryMessages(3, intArrayPass)
				saryTempMessages(4, UBound(saryTempMessages, 2)) = saryMessages(4, intArrayPass)
		
			Next
		
			'Save the changes
			saryMessages = saryTempMessages
		
			'Read in the messages to the application variable
			Application("ChatsarryAppChatMessages") = saryMessages
		
		End If

	End If

	'Unlock the application
	Application.UnLock

End Function



'*********************************
'***    Reset the chat room    ***
'*********************************

Private Function resetFGCWebChat()

	'Dimesion variables
	Dim saryMessages(4, 0)
	Dim saryWebChatUsers(10, 0)

	'Lock the application so that no other user can try and update the application level variable at the same time
	Application.Lock
	
	'Read in the messages to the application variable
	Application("ChatsarryAppChatMessages") = saryMessages

	'Read in the chatroom users to the application variable
	Application("ChatsarryAppChatUsers") = saryWebChatUsers

	'Unlock the application
	Application.UnLock

End Function



'*****************************************
'***    Check for chatroom messages    ***
'*****************************************

Private Function checkMessages()
	
	'Check for messages
	If NOT isArray(Application("ChatsarryAppChatMessages")) Then

		'Dimesion variables
		Dim saryMessages(4, 0)

		'Lock the application so that no other user can try and update the application level variable at the same time
		Application.Lock
	
		'Read in the messages to the application variable
		Application("ChatsarryAppChatMessages") = saryMessages

		'Unlock the application
		Application.UnLock

	End If

End Function



'********************************
'***    Format the message    ***
'********************************

Private Function formatMessage(ByVal strMessage)

	'Dimesion variables
	Dim strTempMessage
	Dim strMessageLink
	Dim lngStartPos
	Dim lngEndPos
	Dim intLoop

	'Loop through the emoticons array
	For intLoop = 1 to UBound(saryEmoticons)
		strMessage = Replace(strMessage, saryEmoticons(intLoop,2), "<img border=""0"" src=""" & saryEmoticons(intLoop,3) & """ title=""" & saryEmoticons(intLoop,1) & """ />", 1, -1, 1)
	Next

	'Loop through the smut array
	For intLoop = 1 to UBound(sarySmut)
		strMessage = Replace(strMessage, sarySmut(intLoop, 1), sarySmut(intLoop, 2), 1, -1, 1)
	Next

	'Change forum codes for bold and italic HTML tags back to the normal satandard HTML tags
	strMessage = Replace(strMessage, "[B]", "<strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/B]", "</strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[I]", "<em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/I]", "</em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[U]", "<u>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/U]", "</u>", 1, -1, 1)

	'Loop through the message till all or any images are turned into HTML images
	Do While InStr(1, strMessage, "[IMG]", 1) > 0  AND InStr(1, strMessage, "[/IMG]", 1) > 0

		'Find the start position in the message of the [IMG] code
		lngStartPos = InStr(1, strMessage, "[IMG]", 1)

		'Find the position in the message for the [/IMG]] closing code
		lngEndPos = InStr(lngStartPos, strMessage, "[/IMG]", 1) + 6
		
		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'Read in the code to be converted into a hyperlink from the message
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))

		'Place the message link into the tempoary message variable
		strTempMessage = strMessageLink

		'Format the IMG tages into an HTML image tag
		strTempMessage = Replace(strTempMessage, "[IMG]", "<img src=""", 1, -1, 1)
		
		'If there is no tag shut off place a > at the end
		If InStr(1, strTempMessage, "[/IMG]", 1) Then
			strTempMessage = Replace(strTempMessage, "[/IMG]", """ onLoad=""checkImageSize(this);"" />", 1, -1, 1)
		Else
			strTempMessage = strTempMessage & " onLoad=""checkImageSize(this);"" />"
		End If

		'Place the new fromatted hyperlink into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)

	Loop

	'Loop through the message till all or any hyperlinks are turned into HTML hyperlinks
	Do While InStr(1, strMessage, "[URL]", 1) > 0  AND InStr(1, strMessage, "[/URL]", 1) > 0

		'Find the start position in the message of the [URL] code
		lngStartPos = InStr(1, strMessage, "[URL]", 1)

		'Find the position in the message for the [/URL]] closing code
		lngEndPos = InStr(lngStartPos, strMessage, "[/URL]", 1) + 6
		
		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'Read in the code to be converted into a hyperlink from the message
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))

		'Place the message link into the tempoary message variable
		strTempMessage = strMessageLink

		'Remove hyperlink BB codes
		strTempMessage = Replace(strTempMessage, "[URL]", "", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, "[/URL]", "", 1, -1, 1)
		
		'Format the URL tages into an HTML hyperlinks
		strTempMessage = "<a href=""" & strTempMessage & """ class=""chatRoomLink"" target=""_blank"">" & strTempMessage & "</a>"
		
		'Place the new fromatted hyperlink into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)

	Loop

	'Loop through the message till all font colour codes are turned into fonts colours
	Do While InStr(1, strMessage, "[COLOR=", 1) > 0  AND InStr(1, strMessage, "[/COLOR]", 1) > 0

		'Find the start position in the message of the [COLOR= code
		lngStartPos = InStr(1, strMessage, "[COLOR=", 1)

		'Find the position in the message for the [/COLOR] closing code
		lngEndPos = InStr(lngStartPos, strMessage, "[/COLOR]", 1) + 8

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 9

		'Read in the code to be converted into a font colour from the message
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))

		'Place the message colour into the tempoary message variable
		strTempMessage = strMessageLink

		'Format the link into an font colour HTML tag
		strTempMessage = Replace(strTempMessage, "[COLOR=", "<font color=", 1, -1, 1)
		'If there is no tag shut off place a > at the end
		If InStr(1, strTempMessage, "[/COLOR]", 1) Then
			strTempMessage = Replace(strTempMessage, "[/COLOR]", "</font>", 1, -1, 1)
			strTempMessage = Replace(strTempMessage, "]", ">", 1, -1, 1)
		Else
			strTempMessage = strTempMessage & ">"
		End If

		'Place the new fromatted colour HTML tag into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)

	Loop

	'Return the function
	formatMessage = strMessage

End Function




'********************************************
'***    Format JavaScript Safe Message    ***
'********************************************

Private Function formatJSMessage(ByVal strMessage)

	'Replace \ with \\
	strMessage = Replace(strMessage, "\", "\\", 1, -1, 1)

	'Replace " with \"
	strMessage = Replace(strMessage, """", "\""", 1, -1, 1)

	'Replace ' with \'
	strMessage = Replace(strMessage, "'", "\'", 1, -1, 1)

	'Replace line with js line
	strMessage = Replace(strMessage, vbCrLf, "\n", 1, -1, 1)

	'Replace tab with js tab
	strMessage = Replace(strMessage, vbTab, "\t", 1, -1, 1)

	'Return the function
	formatJSMessage = strMessage

End Function



'*************************************
'***    Set the last message ID    ***
'*************************************

Public Function setChatSessionID(ByVal lngSessionID)

	'Check for a valid ID
	If isNumeric(lngSessionID) Then
		
		'Set the Session for the last message ID
		Session("ChatLastMessageID") = CLng(lngSessionID)
		
	End If

End Function



'*************************************
'***    Get the last message ID    ***
'*************************************

Public Function getChatSessionID()

	'Check for a valid ID
	If isNumeric(Session("ChatLastMessageID")) Then
		'Set the Session for the last message ID
		getChatSessionID = CLng(Session("ChatLastMessageID"))
	
	'Else return 0 as the ID
	Else

	intLastMessageID = CLng(saryMessages(1, lngLastMessageIndexPos))
		getChatSessionID = CLng(saryMessages(1, lngLastMessageIndexPos))-1
	End If

End Function

%>