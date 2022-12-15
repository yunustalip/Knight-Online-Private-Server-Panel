<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<%
'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension Variables
Dim saryMessages
Dim strUsername
Dim strMessage
Dim intLastMessageID
Dim strPMUsername
Dim lngLastMessageIndexPos
Dim intChatRoomID
Dim blnUserFound
Dim blnNewMessage


'Intialise Variables
blnUserFound = False
blnNewMessage = False


'Creat Objects


'Read in requests
%>
<html>
<head>
<meta http-equiv="refresh" content="<% = intRefreshTimeout %>" />
<title><% = strWebsiteName & " - " & strTxtChatroom %></title>
<script language="JavaScript">
<!--
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
dim msgsira
'Check for new message

If getChatSessionID() <> CLng(saryMessages(1, lngLastMessageIndexPos)) 
for x=1 to CLng(saryMessages(1, lngLastMessageIndexPos))-getChatSessionID()
msgsira=lngLastMessageIndexPos-(CLng(saryMessages(1, lngLastMessageIndexPos))-getChatSessionID())+x
if msgsira<=0  exit for
	'Read in the chat info
	strUsername = saryMessages(0, msgsira)

	'Read in the message
	strMessage = saryMessages(3, msgsira)

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
Response.Write "postMessage('"&Trim(strMessage)&"');"
next
	'Read in the message ID
	intLastMessageID = CLng(saryMessages(1, lngLastMessageIndexPos))

	'Set the Session for the last message ID
	setChatSessionID(intLastMessageID)

End If


'Check if to log out user
If NOT CheckUsername(strLoggedInUsername)  Response.Write("top.location.href = 'logout.asp';")

%>
function postMessage(message) {

	var objMessage = window.parent.getObject("chatBody");
	var sMessage = message;

	//Post message
	if (sMessage == "/clear") {
		objMessage.innerHTML = "";
	} else if (sMessage != "") {
		objMessage.innerHTML += sMessage + "<br /><br />";
	}

	//Scroll to the end of the page
	if (sMessage != "") objMessage.scrollTop = objMessage.scrollHeight;
}


function checkAvatar(sAvatar) {
	if (sAvatar.width > <% = intAvatarWidth %>) sAvatar.style.width = <% = intAvatarWidth %>;
	if (sAvatar.height > <% = intAvatarHeight %>) sAvatar.style.height = <% = intAvatarHeight %>;
}
<%

'Check if to play sound
If blnNewMessage AND blnPlayNewMessageSound 
	Response.Write(vbCrLf & "//Play the sound to nofify the user of a new message")
	Response.Write(vbCrLf & "if (window.parent.getObject(""sound"").value == ""on"") {")
	Response.Write(vbCrLf & "	window.parent.getFlashMovieObject(""newMsgSound"").Play();")
	Response.Write(vbCrLf & "}")
End If

%>
-->
</script>


<!--#include file="includes/header.asp"-->
<%

Response.Write(vbCrLf & "	<table width=""100%"" height=""100%"" cellpadding=""3"" cellspacing=""2"" bgcolor=""" & strTableColour & """>")
Response.Write(vbCrLf & "	<tr>")
Response.Write(vbCrLf & "		<td valign=""top"">")
Response.Write(vbCrLf & "		  <div class=""smText"" style=""overflow: auto; width: 100%; height: 100%;"">")
Response.Write(vbCrLf & "		  <table cellpadding=""3"" cellspacing=""1"" width=""100%"" bgcolor=""" & strTableBorderColour & """>")
Response.Write(vbCrLf & "			<tr><td bgcolor=""" & strTableBgColour2 & """ align=""center"" class=""bold"">" & UBound(saryWebChatUsers, 2) & " " & strTxtUsersInChatroom & "</td></tr>")

'Get the online users
For intArrayPass = 1 To UBound(saryWebChatUsers, 2)

	'Display the users in chatroom
	Response.Write(vbCrLf & "			<tr><td bgcolor=""" & strTableBgColour2 & """><img src=""" & saryWebChatUsers(9, intArrayPass) & """ align=""middle"" border=""0"" onLoad=""checkAvatar(this);"" />&nbsp;")
	
	'If the user is idle
	If DateDiff("s", CDate(saryWebChatUsers(10, intArrayPass)), NOW()) >= intIdelUserTime 
		Response.Write("<span class=""lgText"" style=""color: " & strIdleUserColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span><br />")
	'Else the user is an admin
	ElseIf CBool(saryWebChatUsers(6, intArrayPass)) 
		Response.Write("<span class=""lgText"" style=""color: " & strAdminColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span><br />")
	'Else for self
	ElseIf lngLoggedInUserID = intArrayPass 
		Response.Write("<span class=""lgText"" style=""color: " & strMeColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span><br />")
	'Else normal user
	Else
		Response.Write("<span class=""lgText"" style=""color: " & strUserColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span><br />")
	End If

	Response.Write("<span class=""bold"">" & strTxtStatus & ":&nbsp;</span><span class=""smText"">" & saryWebChatUsers(8, intArrayPass) & "</span><br />")
	If lngLoggedInUserID = intArrayPass OR blnAdmin  Response.Write("<span class=""bold"">" & strTxtIP & ":&nbsp;</span><span class=""smText"">" & saryWebChatUsers(0, intArrayPass) & "</span><br />")
	Response.Write("<span class=""bold"">" & strTxtBrowser & ":&nbsp;</span><span class=""smText"">" & saryWebChatUsers(5, intArrayPass) & "</span><br />")
	Response.Write("<span class=""bold"">" & strTxtDuration & ":&nbsp;</span><span class=""smText"">" & getInterval(DateDiff("n", CDate(saryWebChatUsers(2, intArrayPass)), CDate(saryWebChatUsers(3, intArrayPass)))) & "</span><br />")
	Response.Write(vbCrLf & "</td></tr>")

Next

Response.Write(vbCrLf & "		  </table>")
Response.Write(vbCrLf & "		  </div>")
Response.Write(vbCrLf & "		</td>")
Response.Write(vbCrLf & "	</tr>")
Response.Write(vbCrLf & "	<tr>")
Response.Write(vbCrLf & "		<td height=""20"" align=""center"" class=""bold""><span style=""color: " & strAdminColor & """>" & strTxtAdmin & "</span> | <span style=""color: " & strIdleUserColor & """>" & strTxtIdleUser & "</span> | <span style=""color: " & strMeColor & """>" & strTxtMe & "</span></td>")
Response.Write(vbCrLf & "	</tr>")
Response.Write(vbCrLf & "	</table>")

%>
<!--#include file="includes/footer.asp"-->