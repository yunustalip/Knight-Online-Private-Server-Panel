<%

'Check for messages
Call checkMessages()

'Read in the messages from the application variable array
saryMessages = Application(strFGCAppPrefix & "sarryAppChatMessages")

'Read in the last message pos
lngLastMessageIndexPos = UBound(saryMessages, 2)

'Read in the chatroom
intChatRoomID = CInt(saryMessages(4, lngLastMessageIndexPos))
dim x
'Check for new message

If getChatSessionID() <> CLng(saryMessages(1, lngLastMessageIndexPos)) 

	'Read in the chat info
	strUsername = saryMessages(0, lngLastMessageIndexPos)

	'Read in the message
	strMessage = saryMessages(3, lngLastMessageIndexPos)

	'Read in the message ID
	intLastMessageID = CLng(saryMessages(1, lngLastMessageIndexPos))

	'Set the Session for the last message ID
	setChatSessionID(intLastMessageID)

	'Set to true to play sound
	blnNewMessage = True

	'Check for commands
	If Left(strMessage, 1) = "/" 

		'**************************
		'***    Clear Screen    ***
		'**************************
		If UCase(Trim(strMessage)) = "/CLEAR" 

			'Check the username
			If strLoggedInUsername = saryMessages(0, lngLastMessageIndexPos) 
				
				'Add the PM prefix
				strMessage = "/clear"

				'Set to true
				blnUserFound = True

			End If

			'Delete message if user was not found
			If NOT blnUserFound  strMessage = ""

		End If

	Else

		'If there is a message to post
		If NOT isNothing(strMessage) AND NOT isNothing(strUsername) 
			
			'Add the username to the message
			strMessage = "<span style=""font-weight: bold;"">" & strUsername& ":</span> " & strMessage

		End If

	End If

	'Format the message for JavaScript
	strMessage = formatJSMessage(strMessage)

End If

%>