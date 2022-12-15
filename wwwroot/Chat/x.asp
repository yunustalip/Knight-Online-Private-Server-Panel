<!--#include file="config.asp"-->
<%
saryMessages = Application("ChatsarryAppChatMessages")
'Read in the last message pos
lngLastMessageIndexPos = UBound(saryMessages, 2)

'Read in the chatroom
intChatRoomID = CInt(saryMessages(4, lngLastMessageIndexPos))

'Check for new message

	'Read in the chat info
	strUsername = saryMessages(0, lngLastMessageIndexPos)

	'Read in the message
	strMessage = saryMessages(3, lngLastMessageIndexPos)

	'Read in the message ID
	intLastMessageID = CLng(saryMessages(1, lngLastMessageIndexPos))
xx=application("chatsarryAppChatUsers")

for xx=1 to UBound(saryMessages,2)
for yy=0 to UBound(saryMessages,1)
Response.Write saryMessages(yy,xx)&" - "
next
Response.Write "<br>"
next
 %>