<!--#include file="admin_config.asp"-->
<!--#include file="language_files/language_file_inc.asp" -->
<!--#include file="functions/functions_chatusers.asp"-->
<!--#include file="includes/skin_file.asp" -->
<!--#include file="includes/emoticons.asp"-->
<!--#include file="includes/smut_inc.asp"-->
<!--#include file="includes/chatrooms_inc.asp"--><%
response.charset="iso-8859-9"
'Set the server timeout
Server.ScriptTimeout = 90

'Set the Session timeout
Session.Timeout = 20

'Holds the version for the chatroom you are running
Const strVersion = "2.0"

Const Notice= "<span style=""font-weight:bold;color:#ff0000"">- Chat Odamýza Hoþgeldiniz. Güzel Vakit Geçirmeniz Dileðiyle.</span><br /><br />"

'Holds the amount of seconds to refresh the chatroom (default = 3)
Const intRefreshTimeout =3


'Set the time to remove un-active users (in seconds)
Const intIdelUserTime = 120

Const Floodsaniye=3

'Holds the time to wait before login out idel users (in minutes)
Const intIdelUserLogoutTime = 5


'Set to true if to play sound to notify the user of new message
Const blnPlayNewMessageSound = True


'Database variables
'------------------------------------------------------------------
Dim adoCon 						'Database Connection Variable Object
Dim strCon						'Holds the string to connect to the db
Dim rsCommon					'Holds the configuartion recordset
Dim strSQL						'Holds the SQL query for the database
Dim blnCode						'Set to true


'User variables
'------------------------------------------------------------------
Dim lngLoggedInUserID			'Holds a logged in users ID number
Dim lngLoggedInUserIP			'Holds a logged in users IP Address
Dim strLoggedInUsername			'Holds a logged in users username
Dim intLoggedInUserroom			'Holds a logged in users chatroom
Dim strLoggedInUserStats		'Holds a logged in users status
Dim blnAdmin					'set to true if the user is an admininstrator


'Initialise variables
'------------------------------------------------------------------
lngLoggedInUserID = 0
strLoggedInUsername = strTxtGuest
strLoggedInUserStats = strTxtAvailable
blnAdmin = False
blnCode = True
'------------------------------------------------------------------


'If Check if the user is an admin
'------------------------------------------------------------------
If Session("ChatChatRoomAdmin") = "1" Then blnAdmin = True


'Initialise the user variables
'------------------------------------------------------------------
If NOT isNotLoggedIn Then
	lngLoggedInUserID = intArrayPos
	lngLoggedInUserIP = saryWebChatUsers(0, intArrayPos)
	strLoggedInUsername = saryWebChatUsers(1, intArrayPos)
	intLoggedInUserroom = saryWebChatUsers(7, intArrayPos)
	strLoggedInUserStats = saryWebChatUsers(8, intArrayPos)
End If

%>
<!--#include file="functions/functions_chatroom.asp"-->