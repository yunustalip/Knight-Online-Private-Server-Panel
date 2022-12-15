<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_config.asp" -->
<!--#include file="includes/skin_file.asp" -->
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
Dim objStream
Dim saryWebChatUsers
Dim strImage
Dim intArrayPass


'Intialise Variables
saryWebChatUsers = Application(strFGCAppPrefix & "sarryAppChatUsers")
strImage = Server.MapPath(strImagePath & "offline.gif")


'Check if there is any users in the chatroom
If isArray(saryWebChatUsers) Then

	'Iterate through the array to see if the user is already in the array
	For intArrayPass = 1 To UBound(saryWebChatUsers, 2)

		'Check if the user is admin
		If CBool(saryWebChatUsers(6, intArrayPass)) Then

			'Change the image
			strImage = Server.MapPath(strImagePath & "online.gif")

			'Exit the Loop
			Exit For
		
		End If

	Next

End If


'Creat the stream object
Set objStream = Server.CreateObject("ADODB.Stream")


'Open the streem oject
objStream.Open


'Set the stream object type to binary
objStream.Type = 1


'Load in the image gif
objStream.LoadFromFile strImage


'Set the right response content type for the image
Response.ContentType = "image/gif"


'Display image
Response.BinaryWrite objStream.Read


'Flush the response object
Response.Flush


'Close the object
objStream.Close


'Release the object
Set objStream = Nothing 

%>