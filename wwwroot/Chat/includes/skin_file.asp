<%

'Read in the chat room theme
Dim strChatRoomTheme : strChatRoomTheme = Session(strFGCAppPrefix & "ChatRoomTheme")


'Dimension Variables
Dim strBgColour
Dim strTableTitleColour2


'***********************
'***    Red Theme    ***
'***********************
If strChatRoomTheme = "red" Then

	strBgColour = "#FF7878"				'Forum page colour
	strTableTitleColour2 = "#C80000"	'Colour of second title if more than one title in a table

'************************
'***    Blue Theme    ***
'************************
ElseIf strChatRoomTheme = "blue" Then

	strBgColour = "#24B8FF"				'Forum page colour
	strTableTitleColour2 = "#008ED2"	'Colour of second title if more than one title in a table

'*************************
'***    Green Theme    ***
'*************************
ElseIf strChatRoomTheme = "green" Then

	strBgColour = "#00CC66"				'Forum page colour
	strTableTitleColour2 = "#009F00"	'Colour of second title if more than one title in a table

'**************************
'***    Orange Theme    ***
'**************************
ElseIf strChatRoomTheme = "orange" Then

	strBgColour = "#FFA953"				'Forum page colour
	strTableTitleColour2 = "#DD6F00"	'Colour of second title if more than one title in a table

'*************************
'***    Black Theme    ***
'*************************
ElseIf strChatRoomTheme = "black" Then

	strBgColour = "#646464"				'Forum page colour
	strTableTitleColour2 = "#000000"	'Colour of second title if more than one title in a table
'*************************
'***    Black2 Theme    ***
'*************************
ElseIf strChatRoomTheme = "black-green" Then

	strBgColour = "#646464"				'Forum page colour
	strTableTitleColour2 = "#000000"	'Colour of second title if more than one title in a table

'***************************
'***    Default Theme    ***
'***************************
Else

	strBgColour = "#24B8FF"				'Forum page colour
	strTableTitleColour2 = "#008ED2"	'Colour of second title if more than one title in a table

End If


'Global on Each Page
'---------------------------------------------------------------------------------

Const strBgImage = ""	'Forum bacground image path
Const strTextColour = "#000000"			'Text colour

'Table colours
'---------------------------------------------------------------------------------
Const strTableColour = "#FFFFFF"		'Table colour
Const strTableBgImage = ""			'Table backgroud image path
Const strTableBgColour = "#FFFFFF"		'Table backgroud colour
Const strTableBorderColour = "#C8C8C8"		'Table border colour
Const strTableVariableWidth = "90%"		'Variable table size

Const strTableTitleColour = "#CCCCCC"		'Table title colour
Const strTableTitleBgImage = ""	'Table backgroud image path

Const strTableBgColour2 = "#F0F0F0"		'Table backgroud colour
Const strTableTitleBgImage2 = ""		'Background image path if more than one title in a table

'User Info
'---------------------------------------------------------------------------------
Const strAdminColor = "#FF0000"
Const strUserColor = "#000000"
Const strMeColor = "#0000FF"
Const strIdleUserColor = "#969696"
Const strFGCWebChatGreen = "#006600"
Const strFGCWebChatRed = "#FF3333"

'Misc
'---------------------------------------------------------------------------------
Const strImagePath = "images/"	'Path to the forum images
Const intAvatarWidth = 35
Const intAvatarHeight = 35
Const intImageWidth = 150
Const intImageHeight = 150

%>