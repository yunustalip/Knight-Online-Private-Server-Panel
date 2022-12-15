
<%if Session("login")<>"ok" Then
Response.Write("<script>alert('Chat Odasina Girmeden Önce Anasayfadan Giris Yapiniz.');location.href ='../';</script>")
Response.End
End If
'Dimension variables
Dim saryWebChatUsers
Dim intArrayPass
Dim intArrayPos
Dim isNotLoggedIn


'Set the user as not logged in
isNotLoggedIn = True


'Array dimension lookup table
' 0 = IP
' 1 = Username
' 2 = Login Time
' 3 = Last Active Time
' 4 = OS
' 5 = Browser
' 6 = Admin
' 7 = Chatroom
' 8 = Status
' 9 = Avatar
' 10 = Last post time
'Application("ChatsarryAppChatUsers") = null

'******************************
'***    Initialise array    ***
'******************************
	
'Initialise  the array from the application veriable
If IsArray(Application("ChatsarryAppChatUsers")) Then 
	
	'Place the application level array into a temporary dynaimic array
	saryWebChatUsers = Application("ChatsarryAppChatUsers")

'Else Initialise the variable as an empty array
Else
	ReDim saryWebChatUsers(10, 0)
End If



'**************************************
'***    Get users array position    ***
'**************************************

'Iterate through the array to see if the user is already in the array
For intArrayPass = 1 To UBound(saryWebChatUsers, 2)

	'Check the IP address and username
	If saryWebChatUsers(0, intArrayPass) = getIP() AND saryWebChatUsers(1, intArrayPass) = Session("ChatChatRoomUsername") Then
		
		intArrayPos = intArrayPass
		isNotLoggedIn = False
	
	End If

Next 


'If the user is found in the array update the array
If NOT isNotLoggedIn Then

	saryWebChatUsers(0, intArrayPos) = getIP()
	saryWebChatUsers(3, intArrayPos) = CDbl(NOW())
	saryWebChatUsers(4, intArrayPos) = OSType()
	saryWebChatUsers(5, intArrayPos) = BrowserType()

	'Update the chatroom
	Call updateChatroom()

End If

'****************************************
'***    Check if user is logged in    ***
'****************************************

Public Function CheckUsername(ByVal strUsername)

	'Dimension variables
	Dim blnUsername

	'Initialise variables
	blnUsername = False

	'Loop through the array of users
	For intArrayPass = 1 To UBound(saryWebChatUsers, 2)
		
		'Check if the username and ID is not already in use
		If saryWebChatUsers(1, intArrayPass) = strUsername Then
			
			'If in use Then initialise the variables to true
			blnUsername = True

			'Exit the loop
			Exit For

		End If

	Next

	'Return the function
	CheckUsername = blnUsername

End Function



'****************************
'***    Add a new user    ***
'****************************

Public Function AddUser(ByVal strUsername, ByVal intChatRoom, ByVal strStatus, ByVal strAvatar)

	'ReDimesion the array
	ReDim Preserve saryWebChatUsers(10, UBound(saryWebChatUsers, 2) + 1)
	
	'Update the new array position which will be the last one
	saryWebChatUsers(0, UBound(saryWebChatUsers, 2)) = getIP()
	saryWebChatUsers(1, UBound(saryWebChatUsers, 2)) = Trim(strUsername)
	saryWebChatUsers(2, UBound(saryWebChatUsers, 2)) = CDbl(NOW())
	saryWebChatUsers(3, UBound(saryWebChatUsers, 2)) = CDbl(NOW())
	saryWebChatUsers(4, UBound(saryWebChatUsers, 2)) = OSType()
	saryWebChatUsers(5, UBound(saryWebChatUsers, 2)) = BrowserType()
	saryWebChatUsers(7, UBound(saryWebChatUsers, 2)) = intChatRoom
	saryWebChatUsers(8, UBound(saryWebChatUsers, 2)) = strStatus
	saryWebChatUsers(9, UBound(saryWebChatUsers, 2)) = strAvatar
	saryWebChatUsers(10, UBound(saryWebChatUsers, 2)) = CDbl(NOW())

	'Update the chatroom users
	Call removeUnActiveUser()

	'Post a message sayin a user has joined the chat room
	Call postMessage("", strUsername & " " & strTxtHasJoinedTheChatroom, "", "", "")

End Function

	

'************************************
'***    Remove un-active users    ***
'************************************

Public Function removeUnActiveUser()

	'Iterate through the array to remove old entires
	For intArrayPass = 1 To UBound(saryWebChatUsers, 2)

		'Check the last checked date if user wasnt updated x seconds ago
		If CDate(saryWebChatUsers(3, intArrayPass)) < DateAdd("n", - intIdelUserLogoutTime, NOW()) Then

		'Post message that the user has left
		Call postMessage("", saryWebChatUsers(1, intArrayPass) & " " & strTxtHasLeftTheChatroom, "", "", "")

			'Swap this array postion with the last in the array
			saryWebChatUsers(0, intArrayPass) = saryWebChatUsers(0, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(1, intArrayPass) = saryWebChatUsers(1, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(2, intArrayPass) = saryWebChatUsers(2, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(3, intArrayPass) = saryWebChatUsers(3, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(4, intArrayPass) = saryWebChatUsers(4, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(5, intArrayPass) = saryWebChatUsers(5, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(6, intArrayPass) = saryWebChatUsers(6, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(7, intArrayPass) = saryWebChatUsers(7, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(8, intArrayPass) = saryWebChatUsers(8, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(9, intArrayPass) = saryWebChatUsers(9, UBound(saryWebChatUsers, 2))
			saryWebChatUsers(10, intArrayPass) = saryWebChatUsers(10, UBound(saryWebChatUsers, 2))

			'Remove the last array position as it is no-longer needed
			ReDim Preserve saryWebChatUsers(10, UBound(saryWebChatUsers, 2) - 1)

			'Exit the loop
			Exit For

		End If

	Next

	'Update the chatroom
	Call updateChatroom()

End Function



'************************************
'***    Logout of the chatroom    ***
'************************************

Public Function logoutUser(ByVal strUsername)

	'Loop through the array of users
	For intArrayPass = 1 To UBound(saryWebChatUsers, 2)
		
		'Check for username
		If saryWebChatUsers(1, intArrayPass) = strUsername Then
			
			'Set the users last active time back to logout
			saryWebChatUsers(3, intArrayPass) = CDbl(DateAdd("n", - (intIdelUserLogoutTime + 10), NOW()))

			'Exit the loop
			Exit For

		End If

	Next

	'Remove the user
	Call removeUnActiveUser()

End Function



'*****************************
'***    Format Username    ***
'*****************************

Private Function formatUsername(ByVal strUsername)

	'Remove hazardous characters
	strUsername = MultiReplace(strUsername, "<>'""/\&`~+=", "")

	'Remove bad words
	For intLoop = 1 to UBound(sarySmut)
		strUsername = Replace(strUsername, sarySmut(intLoop, 1), sarySmut(intLoop, 2), 1, -1, 1)
	Next

	'Return the function
	formatUsername = strUsername

End Function



'*************************************
'***    Change the users status    ***
'*************************************

Public Function changeStatus(ByVal strNewStatus)

	'Change the users status
	saryWebChatUsers(8, intArrayPos) = strNewStatus

	'Update the chatroom
	Call updateChatroom()

	'Post message that the user has left
	Call postMessage("", saryWebChatUsers(1, intArrayPos) & ", " & strTxtStatusIsNow & " [B]" & strNewStatus & "[/B] Olarak Deðiþtirdi.", "", "", "")

End Function



'*****************************
'***    Change Username    ***
'*****************************

Public Function changeUsername(ByVal strUsername, ByVal strNewUsername)

	'Format the new username
	strNewUsername = formatUsername(strNewUsername)

	'Trim the username down to 20 characters
	strNewUsername = Mid(strNewUsername, 1, 20)

	'Change the users status
	saryWebChatUsers(1, intArrayPos) = strNewUsername

	'Change the username in the username Session
	Session("ChatChatRoomUsername") = strNewUsername

	'Update the chatroom
	Call updateChatroom()

	'Post a message that the username is changed
	Call postMessage("", strUsername & " " & strTxtIsNowKnownAs & " [B]" & strNewUsername & "[/B]", "", strFGCWebChatGreen, "")

End Function


'****************************
'***    Login in admin    ***
'****************************

Private Function loginAdmin(ByVal strPassword)

	'Check if the password match
	If Trim(strPassword) = strChatRoomAdminPassword Then
		
		'Create the admin Session
		Session("ChatChatRoomAdmin") = "1"

		'Set the user to admin
		saryWebChatUsers(6, intArrayPos) = True

		'Update the chatroom
		Call updateChatroom()

		'Redirect the page for settings to take effect
		Response.Write "<script language=""JavaScript"">alert(""Admin girisi yapildi!"");</script>"

	End If

End Function



'***************************************
'***    Update the chatroom users    ***
'***************************************

Public Function updateChatroom()

	'Lock the application so that no other user can try and update the application level variable at the same time
	Application.Lock

	'Update the application level variable
	Application("ChatsarryAppChatUsers") = saryWebChatUsers

	'Unlock the application
	Application.UnLock

End Function

Private Function BrowserType()

	Dim strUserAgent	'Holds info on the users browser and os
	Dim strBrowserUserType	'Holds the users browser type

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Get the uesrs web browser
	'Opera
	If InStr(1, strUserAgent, "Opera", 1) > 0 Then
		strBrowserUserType = "Opera"

	'AOL
	ElseIf inStr(1, strUserAgent, "AOL", 1) > 0  Then
		strBrowserUserType = "AOL"

	'Konqueror
	ElseIf inStr(1, strUserAgent, "Konqueror", 1) > 0 Then
		strBrowserUserType = "Konqueror"

	'EudoraWeb
	ElseIf inStr(1, strUserAgent, "EudoraWeb", 1) > 0 Then
		strBrowserUserType = "EudoraWeb"

	'Dreamcast
	ElseIf inStr(1, strUserAgent, "Dreamcast", 1) > 0 Then
		strBrowserUserType = "Dreamcast"
	
	'Safari
	ElseIf inStr(1, strUserAgent, "Safari", 1) > 0 Then
		strBrowserUserType = "Safari"
	
	'Lynx
	ElseIf inStr(1, strUserAgent, "Lynx", 1) > 0 Then
		strBrowserUserType = "Lynx"
	
	'ICE
	ElseIf inStr(1, strUserAgent, "ICE", 1) > 0 Then
		strBrowserUserType = "ICE"
	
	'iCab 
	ElseIf inStr(1, strUserAgent, "iCab", 1) > 0 Then
		strBrowserUserType = "iCab"
		
	'HotJava 
	ElseIf inStr(1, strUserAgent, "Sun", 1) > 0 AND inStr(1, strUserAgent, "Mozilla/3", 1) > 0 Then
		strBrowserUserType = "HotJava"
	
	'Galeon 
	ElseIf inStr(1, strUserAgent, "Galeon", 1) > 0 Then
		strBrowserUserType = "Galeon"
		
	'Epiphany 
	ElseIf inStr(1, strUserAgent, "Epiphany", 1) > 0 Then
		strBrowserUserType = "Epiphany"
	
	'DocZilla 
	ElseIf inStr(1, strUserAgent, "DocZilla", 1) > 0 Then
		strBrowserUserType = "DocZilla"
	
	'Camino 
	ElseIf inStr(1, strUserAgent, "Chimera", 1) > 0 OR inStr(1, strUserAgent, "Camino", 1) > 0 Then
		strBrowserUserType = "Camino"
	
	'Dillo 
	ElseIf inStr(1, strUserAgent, "Dillo", 1) > 0 Then
		strBrowserUserType = "Dillo"
		
	'amaya 
	ElseIf inStr(1, strUserAgent, "amaya", 1) > 0 Then
		strBrowserUserType = "Amaya"
		
	'NetCaptor 
	ElseIf inStr(1, strUserAgent, "NetCaptor", 1) > 0 Then
		strBrowserUserType = "NetCaptor"
		
		
		
	'Internet Explorer
	ElseIf inStr(1, strUserAgent, "MSIE", 1) > 0 Then
		strBrowserUserType = "Microsoft IE"

	'Pocket Internet Explorer
	ElseIf inStr(1, strUserAgent, "MSPIE", 1) > 0 Then
		strBrowserUserType = "Pocket IE"

	
	'Mozilla Firefox
	ElseIf inStr(1, strUserAgent, "Gecko", 1) > 0 AND inStr(1, strUserAgent, "Firefox", 1) > 0 Then
		strBrowserUserType = "Firefox"
	
	'Mozilla Firebird
	ElseIf inStr(1, strUserAgent, "Gecko", 1) > 0 AND inStr(1, strUserAgent, "Firebird", 1) > 0 Then
		strBrowserUserType = "Firebird"
	
	'Mozilla
	ElseIf inStr(1, strUserAgent, "Gecko", 1) > 0 AND inStr(1, strUserAgent, "Netscape", 1) = 0 Then
		strBrowserUserType = "Mozilla"

	'Netscape
	ElseIf inStr(1, strUserAgent, "Netscape/", 1) > 0 Then
		strBrowserUserType = "Netscape"

		
	'Else unknown or robot
	Else
		strBrowserUserType = "Unknown"
	End If

	'Return function
	BrowserType = strBrowserUserType

End Function


'******************************************
'***          Get OS Type   	      *****
'******************************************

Private Function OSType()

	Dim strUserAgent	'Holds info on the users browser and os
	Dim strOS		'Holds the users OS

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Get users OS
	'Windows
	If inStr(1, strUserAgent, "Windows 2003", 1) > 0 Or inStr(1, strUserAgent, "NT 5.2", 1) > 0 Then
		strOS = "Windows 2003"
	ElseIf inStr(1, strUserAgent, "Windows XP", 1) > 0 Or inStr(1, strUserAgent, "NT 5.1", 1) > 0 Then
		strOS = "Windows XP"
	ElseIf inStr(1, strUserAgent, "Windows 2000", 1) > 0 Or inStr(1, strUserAgent, "NT 5", 1) > 0 Then
		strOS = "Windows 2000"
	ElseIf inStr(1, strUserAgent, "Windows NT", 1) > 0 Or inStr(1, strUserAgent, "WinNT", 1) > 0 Then
		strOS = "Windows  NT 4"
	ElseIf inStr(1, strUserAgent, "Windows 95", 1) > 0 Or inStr(1, strUserAgent, "Win95", 1) > 0 Then
		strOS = "Windows 95"
	ElseIf inStr(1, strUserAgent, "Windows ME", 1) > 0 Or inStr(1, strUserAgent, "Win 9x 4.90", 1) > 0 Then
		strOS = "Windows ME"
	ElseIf inStr(1, strUserAgent, "Windows 98", 1) > 0 Or inStr(1, strUserAgent, "Win98", 1) > 0 Then
		strOS = "Windows 98"
	ElseIf Instr(1, strUserAgent, "Windows 3.1", 1) > 0 or Instr(1, strUserAgent, "Win16", 1) > 0 Then
		strOS = "Windows 3.x"
	ElseIf Instr(1, strUserAgent, "Windows CE", 1) > 0 Then
		strOS = "Windows CE"

	'PalmOS
	ElseIf inStr(1, strUserAgent, "PalmOS", 1) > 0 Then
		strOS = "Palm OS"
		
	'PalmPilot
	ElseIf inStr(1, strUserAgent, "Elaine", 1) > 0 Then
		strOS = "PalmPilot"

	'Nokia
	ElseIf inStr(1, strUserAgent, "Nokia", 1) > 0 Then
		strOS = "Nokia"

	'Linux
	ElseIf inStr(1, strUserAgent, "Linux", 1) > 0 Then
		strOS = "Linux"

	'Amiga
	ElseIf inStr(1, strUserAgent, "Amiga", 1) > 0 Then
		strOS = "Amiga"

	'Solaris
	ElseIf inStr(1, strUserAgent, "Solaris", 1) > 0 Then
		strOS = "Solaris"

	'SunOS
	ElseIf inStr(1, strUserAgent, "SunOS", 1) > 0 Then
		strOS = "Sun OS"

	'BSD
	ElseIf inStr(1, strUserAgent, "BSD", 1) > 0 or inStr(1, strUserAgent, "FreeBSD", 1) > 0 Then
		strOS = "Free BSD"

	'Unix
	ElseIf inStr(1, strUserAgent, "Unix", 1) > 0 OR inStr(1, strUserAgent, "X11", 1) > 0 Then
		strOS = "Unix"

	'AOL webTV
	ElseIf inStr(1, strUserAgent, "AOLTV", 1) > 0 OR inStr(1, strUserAgent, "AOL_TV", 1) > 0 Then
		strOS = "AOL TV"
	ElseIf inStr(1, strUserAgent, "WebTV", 1) > 0 Then
		strOS = "Web TV"

	'Machintosh
	ElseIf inStr(1, strUserAgent, "Mac OS X", 1) > 0 Then
		strOS = "Mac OS X"
	ElseIf inStr(1, strUserAgent, "Mac_PowerPC", 1) > 0 or Instr(1, strUserAgent, "PPC", 1) > 0 Then
		strOS = "Mac PowerPC"
	ElseIf (inStr(1, strUserAgent, "6800", 1) > 0 OR inStr(1, strUserAgent, "68k", 1) > 0) AND inStr(1, strUserAgent, "Mac", 1) > 0 Then
		strOS = "Mac 68k"
	ElseIf inStr(1, strUserAgent, "Mac", 1) > 0 or inStr(1, strUserAgent, "apple", 1) > 0 Then
		strOS = "Macintosh"

	'OS/2
	ElseIf inStr(1, strUserAgent, "OS/2", 1) > 0 Then
		strOS = "OS/2"


	'Search Robot
	ElseIf inStr(1, strUserAgent, "Googlebot", 1) > 0 OR inStr(1, strUserAgent, "ZyBorg", 1) > 0 OR inStr(1, strUserAgent, "slurp", 1) > 0 OR inStr(1, strUserAgent, "Scooter", 1) > 0 OR inStr(1, strUserAgent, "Robozilla", 1) > 0 OR inStr(1, strUserAgent, "Ask Jeeves", 1) > 0 OR inStr(1, strUserAgent, "Ask+Jeeves", 1) > 0 OR inStr(1, strUserAgent, "lycos", 1) > 0 OR inStr(1, strUserAgent, "ArchitextSpider", 1) > 0 OR inStr(1, strUserAgent, "Gulliver", 1) > 0 OR inStr(1, strUserAgent, "crawler@fast", 1) > 0 Then
		strOS = "Search Robot"
		
	'Search Robot
	ElseIf inStr(1, strUserAgent, "TurnitinBot", 1) > 0 OR inStr(1, strUserAgent, "internetseer", 1) > 0 OR inStr(1, strUserAgent, "nameprotect", 1) > 0 OR inStr(1, strUserAgent, "PhpDig", 1) > 0 OR inStr(1, strUserAgent, "StackRambler", 1) > 0 OR inStr(1, strUserAgent, "UbiCrawler", 1) > 0 OR inStr(1, strUserAgent, "Speedy+Spider", 1) > 0 OR inStr(1, strUserAgent, "ia_archiver", 1) > 0 OR inStr(1, strUserAgent, "msnbot", 1) > 0 OR inStr(1, strUserAgent, "arianna.libero.it", 1) > 0 Then
		strOS = "Search Robot"		
 

	Else
		strOS = "Unknown"
	End If

	'Return function
	OSType = strOS
End Function




'******************************************
'***	 Get users IP address		***
'******************************************

Private Function getIP()

	Dim strIPAddr

	'If they are not going through a proxy get the IP address
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
		strIPAddr = Request.ServerVariables("REMOTE_ADDR")
	
	'If they are going through multiple proxy servers only get the fisrt IP address in the list (,)
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
		strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
	
	'If they are going through multiple proxy servers only get the fisrt IP address in the list (;)
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
		strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
		
	'Get the browsers IP address not the proxy servers IP
	Else
		strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	End If
	
	'Remove all tags in IP string
	strIPAddr =  removeTags(strIPAddr)
	
	'Place the IP address back into the returning function
	getIP = Trim(strIPAddr)

End Function




'******************************************
'***  		Format user input     *****
'******************************************

'Format user input function
Private Function formatInput(ByVal strInputEntry)

	'Get rid of malicous code in the message
	strInputEntry = Replace(strInputEntry, "</script>", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<script language=""javascript"">", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<script language=javascript>", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "script", "&#115;cript", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Script", "&#083;cript", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "script", "&#083;cript", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "object", "&#111;bject", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "OBJECT", "&#079;BJECT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Object", "&#079;bject", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "object", "&#079;bject", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "applet", "&#097;pplet", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "APPLET", "&#065;PPLET", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Applet", "&#065;pplet", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "applet", "&#065;pplet", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "embed", "&#101;mbed", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "EMBED", "&#069;MBED", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Embed", "&#069;mbed", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "embed", "&#069;mbed", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "event", "&#101;vent", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "EVENT", "&#069;VENT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Event", "&#069;vent", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "event", "&#069;vent", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "document", "&#100;ocument", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "DOCUMENT", "&#068;OCUMENT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Document", "&#068;ocument", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "document", "&#068;ocument", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "cookie", "&#099;ookie", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "COOKIE", "&#067;OOKIE", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Cookie", "&#067;ookie", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "cookie", "&#067;ookie", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "form", "&#102;orm", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "FORM", "&#070;ORM", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Form", "&#070;orm", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "form", "&#070;orm", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "iframe", "i&#102;rame", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "IFRAME", "I&#070;RAME", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Iframe", "I&#102;rame", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "iframe", "i&#102;rame", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "textarea", "&#116;extarea", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "TEXTAREA", "&#84;EXTAREA", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Textarea", "&#84;extarea", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "textarea", "&#84;extarea", 1, -1, 1)

	'Return the function
	formatInput = strInputEntry

End Function



'****************************
'***    Strip all tags    ***
'****************************

'Remove all tags for text only display
Private Function removeTags(ByVal strInputEntry)

	'Remove all HTML scripting tags etc. for plain text output
	strInputEntry = Replace(strInputEntry, "&", "&amp;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "&lt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "&gt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "'", "&#146;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, """", "&quot;", 1, -1, 1)

	'Return the function
	removeTags = strInputEntry

End Function



'*****************************************
'***    If the variable has a value    ***
'*****************************************

Private Function isNothing(strVariable)

	'If no value found Then return true
	If isNull(strVariable) OR isEmpty(strVariable) OR Len(strVariable) < 1 Then		
		isNothing = True
	'Else return true
	Else		
		isNothing = False
	End If

End Function



'***********************************
'***    Clean Up And Redirect    ***
'***********************************

Public Function Redirect(ByVal strLocation)

	'Redirect
	Response.Redirect(strLocation)
	Response.End

End Function



'************************************
'***    Multi Replace a String    ***
'************************************

Private Function MultiReplace(ByVal strInputString, ByVal strTxtToReplace, ByVal strTxtToReplaceWith)

	'Dimension Variables
	Dim intLoop
	Dim strReplaceChar

	'Loop through the string of un-wanted characters and replace them
	For intLoop = 1 To Len(strTxtToReplace)

		'Shorten the string to a character
		strReplaceChar = Mid(strTxtToReplace, intLoop, 1)

		'Replace the unwanted string with the new string
		strInputString = Replace(strInputString, strReplaceChar, strTxtToReplaceWith, 1, -1, 1)

	Next

	'Return the function
	MultiReplace = strInputString

End Function



'**************************************
'***    Format the time interval    ***
'**************************************

Private Function getInterval(ByVal lngInterval)

	'If the user has been active for less than 1 hour
	If lngInterval < 60 Then
		getInterval = lngInterval & " Dakika"
	
	'If the user has been active for less than 2 hours
	ElseIf lngInterval < 120 Then
		getInterval = Int(lngInterval / 60) & " Saat " & (lngInterval MOD 60) & " Dakika"
	
	'If the user has been active for less than 24 hours
	ElseIf lngInterval < 1440 Then
		getInterval = Int(lngInterval / 60) & " Saat " & (lngInterval MOD 60) & " Dakika"
	
	'If the user has been active for less than 48 hours
	ElseIf lngInterval < 2880 Then
		getInterval = Int(lngInterval / 24) & " day " & Int(lngInterval / 60) & " Saat " & (lngInterval MOD 60) & " Dakika"
	
	'If the user has been active for less than 30 days
	ElseIf lngInterval < 36000 Then
		getInterval = Int(lngInterval / 24) & " days " & Int(lngInterval / 60) & " Saat " & (lngInterval MOD 60) & " Dakika"
	
	'Else display the active user time in minutes
	Else
		getInterval = lngInterval & " Dakika"
	End If

End Function
%>