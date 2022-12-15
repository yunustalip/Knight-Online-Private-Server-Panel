<%
<!--#include file="../includes/skin_file.asp" -->
Const strFGCAppPrefix = "Chat"
saryWebChatUsers = Application(strFGCAppPrefix & "sarryAppChatUsers")

saryMessages = Application(strFGCAppPrefix & "sarryAppChatMessages")
	If NOT isArray(saryMessages) Then
		
		'ReDimesion the array
		ReDim saryMessages(4, 0)

		'Read in the messages to the application variable
		Application("ChatsarryAppChatMessages") = saryMessages

	End If

for yy=1 to ubound(saryWebChatUsers,2)
for xx=0 to ubound(saryWebChatUsers)
Response.Write saryWebChatUsers(xx,yy)&"-"
next
Response.Write "<br>"
next
%>