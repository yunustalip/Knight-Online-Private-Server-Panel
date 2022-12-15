<!--#include file="../_inc/conn.asp"-->
<!--#include file="../guvenlik.asp"-->
<!--#include file="config.asp" -->
<%
response.charset="iso-8859-9"
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

If NOT CheckUsername(strLoggedInUsername) Then
Session("ChatChatRoomAdmin") = ""
Session("ChatRoomAdmin") = NULL
Session("ChatChatRoomUsername") = ""
Session("ChatChatRoomUsername") = NULL
Session("ChatLastMessageID") = ""
Session("ChatLastMessageID") = NULL
Response.Write("<script>alert('Odadan Atýldýnýz.');location.href ='default.asp';</script>")
Response.End
End If
'Creat Objects


'Read in requests
%><base href="http://<%=Request.ServerVariables("Server_Name")%>/Chat/">
<script language="JavaScript">
<!--
<%
'Check if to log out user

'Check for messages
Call checkMessages()

'Read in the messages from the application variable array
saryMessages = Application("ChatsarryAppChatMessages")

'Read in the last message pos
lngLastMessageIndexPos = UBound(saryMessages, 2)

'Read in the chatroom
intChatRoomID = CInt(saryMessages(4, lngLastMessageIndexPos))
dim x
dim msgsira
'Check for new message

If getChatSessionID() <> CLng(saryMessages(1, lngLastMessageIndexPos)) Then
for x=1 to CLng(saryMessages(1, lngLastMessageIndexPos))-getChatSessionID()
msgsira=lngLastMessageIndexPos-(CLng(saryMessages(1, lngLastMessageIndexPos))-getChatSessionID())+x
if msgsira<=0 Then exit for
	'Read in the chat info
	strUsername = saryMessages(0, msgsira)

	'Read in the message
	strMessage = saryMessages(3, msgsira)

	'Set to true to play sound
	blnNewMessage = True

	'Check for commands
	If Left(strMessage, 1) = "/" Then

		'**************************
		'***    Clear Screen    ***
		'**************************
		If UCase(Trim(strMessage)) = "/CLEAR" Then

			'Check the username
			If strLoggedInUsername = saryMessages(0, lngLastMessageIndexPos) Then
				
				'Add the PM prefix
				strMessage = "/clear"

				'Set to true
				blnUserFound = True

			End If

			'Delete message if user was not found
			If NOT blnUserFound Then strMessage = ""

		End If

	Else

		'If there is a message to post
		If NOT isNothing(strMessage) AND NOT isNothing(strUsername) Then
			If blnAdmin=True Then
			strUsername="<span style=""font-weight: bold;color:red"">" & strUsername& ":</span> "
			Else
			strUsername="<span style=""font-weight: bold;color:black"">" & strUsername& ":</span> "
			End If
			'Add the username to the message
			strMessage = strUsername & strMessage & " ("&hour(cdate(saryMessages(2, msgsira)))&":"&minute(cdate(saryMessages(2, msgsira)))&":"&second(cdate(saryMessages(2, msgsira)))&")"

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



%>
function postMessage(message) {
	var objMessage = document.getElementById("chatBody");
	var sMessage = message;

	//Post message
	if (sMessage == "/clear") {
		objMessage.innerHTML = "";
	} else if (sMessage != "") {
		$('#chatBody').append(sMessage + '<br /><br />');
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
If blnNewMessage AND blnPlayNewMessageSound Then
	Response.Write(vbCrLf & "//Play the sound to nofify the user of a new message")
	Response.Write(vbCrLf & "if (window.parent.getObject(""sound"").value == ""on"") {")
	Response.Write(vbCrLf & "	window.parent.getFlashMovieObject(""newMsgSound"").Play();")
	Response.Write(vbCrLf & "}")
End If

%>
-->
</script>
<%
Response.Write(vbCrLf & "	<table width=""100%"" height=""100%"" cellpadding=""3"" cellspacing=""2"" bgcolor=""" & strTableColour & """>")
Response.Write(vbCrLf & "	<tr>")
Response.Write(vbCrLf & "		<td valign=""top"">")
Response.Write(vbCrLf & "		  <div class=""smText"" style=""width: 100%; height: 100%;"">")
Response.Write(vbCrLf & "		  <table cellpadding=""3"" cellspacing=""1"" width=""100%"" bgcolor=""" & strTableBorderColour & """>")
Response.Write(vbCrLf & "			<tr><td bgcolor=""" & strTableBgColour2 & """ align=""center"" class=""bold"">" & UBound(saryWebChatUsers, 2) & " " & strTxtUsersInChatroom & "</td></tr>")

'Get the online users
For intArrayPass = 1 To UBound(saryWebChatUsers, 2)

'Display the users in chatroom
	Response.Write(vbCrLf & "			<tr><td bgcolor=""" & strTableBgColour2 & """><img src=""" & saryWebChatUsers(9, intArrayPass) & """ align=""middle"" border=""0"" onLoad=""checkAvatar(this);"" />&nbsp;<a href=""javascript:launchChat2('../Karakter-Detay/"&saryWebChatUsers(1, intArrayPass)&"')"">")
	}
	'If the user is idle
	If DateDiff("s", CDate(saryWebChatUsers(10, intArrayPass)), NOW()) >= intIdelUserTime Then
		Response.Write("<span class=""lgText"" style=""color: " & strIdleUserColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span>")
	'Else the user is an admin
	ElseIf CBool(saryWebChatUsers(6, intArrayPass)) Then
		Response.Write("<span class=""lgText"" style=""color: " & strAdminColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span>")
	'Else for self
	ElseIf lngLoggedInUserID = intArrayPass Then
		Response.Write("<span class=""lgText"" style=""color: " & strMeColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span>")
	'Else normal user
	Else
		Response.Write("<span class=""lgText"" style=""color: " & strUserColor & """>" & saryWebChatUsers(1, intArrayPass) & "</span>")
	End If
	Response.Write("</a>")
	If blnAdmin Then
	Response.Write("&nbsp;(<a href=""#"" onclick=""document.frmFGCWebChat.message.value='/ban "&saryWebChatUsers(1, intArrayPass)&"';return false"">Banla</a> - <a href=""#"" onclick=""document.frmFGCWebChat.message.value='/kick "&saryWebChatUsers(1, intArrayPass)&"';return false"">Kickle</a>)")
	End If
	Response.Write("<br /><span class=""bold"">" & strTxtStatus & ":&nbsp;</span><span class=""smText"">" & saryWebChatUsers(8, intArrayPass) & "</span><br />")
	If blnAdmin Then Response.Write("<span class=""bold"">" & strTxtIP & ":&nbsp;</span><span class=""smText"">" & saryWebChatUsers(0, intArrayPass) & "</span><br />")
	If blnAdmin Then Response.Write("<span class=""bold"">" & strTxtBrowser & ":&nbsp;</span><span class=""smText"">" & saryWebChatUsers(5, intArrayPass) & "</span><br />")
	Response.Write("<span class=""bold"">" & strTxtDuration & ":&nbsp;</span><span class=""smText"">" & getInterval(DateDiff("n", CDate(saryWebChatUsers(2, intArrayPass)), CDate(saryWebChatUsers(3, intArrayPass)))) & "</span><br />")
	Response.Write(vbCrLf & "</td></tr>")

Next

Response.Write(vbCrLf & "	  </table>")
Response.Write(vbCrLf & "	  </div>")
Response.Write(vbCrLf & "	</td>")
Response.Write(vbCrLf & "	</tr>")
Response.Write(vbCrLf & "	<tr>")
Response.Write(vbCrLf & "	<td height=""20"" align=""center"" class=""bold""><span style=""color: " & strAdminColor & """>" & strTxtAdmin & "</span> | <span style=""color: " & strIdleUserColor & """>" & strTxtIdleUser & "</span> | <span style=""color: " & strMeColor & """>" & strTxtMe & "</span></td>")
Response.Write(vbCrLf & "	</tr>")
Response.Write(vbCrLf & "	</table>")

%>