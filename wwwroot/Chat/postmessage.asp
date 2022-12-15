<% @ Language=VBScript %>
<% Option Explicit
response.charset="iso-8859-9"%>

<!--#include file="config.asp" -->
<%
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension Variables
Dim strMode
Dim strStatus
Dim strMessage
Dim strFont
Dim strColor
Dim strFormat
Dim strUsername
Dim strPassword
Dim intChatroom
Dim intLoop
Dim nesne
Dim yaz
Dim ag
Dim intArrayNo
Dim SonMesaj

if isNotLoggedIn=true Then
Response.Write("<script>location.href ='default.asp';</script>")
Response.End
End If

'Intialise Variables
SonMesaj=Datediff("s",Cdate(saryWebChatUsers(10, lngLoggedInUserID)),now())
If SonMesaj<=Floodsaniye AND Not blnAdmin Then
Response.Write("<script>alert('Flood Yapmayýnýz. "&Floodsaniye+1&" Saniyede Bir Mesaj Yazabilirsiniz.\nKalan Saniye: "&Floodsaniye-SonMesaj+1&"');</script>")
Response.End
End If

'Creat Objects
Function ConvertFromUTF8(sIn)
Dim oIn: Set oIn = CreateObject("ADODB.Stream")

oIn.Open
oIn.CharSet = "WIndows-1254"
oIn.WriteText sIn
oIn.Position = 0
oIn.CharSet = "UTF-8"
ConvertFromUTF8 = oIn.ReadText
oIn.Close

End Function

'Read in requests
strMode = ConvertFromUTF8(Request.Form("mode"))
strStatus = ConvertFromUTF8(Request.Form("status"))
strMessage = ConvertFromUTF8(Trim(Request.Form("message")))
strFont = ConvertFromUTF8(Trim(Request.Form("font")))
strColor = ConvertFromUTF8(Trim(Request.Form("color")))
strFormat = ConvertFromUTF8(Trim(Request.Form("format")))
intChatroom = ConvertFromUTF8(CInt(Request.Form("chatroom")))



'Check for commands
If Left(strMessage, 6) = "/admin" OR Trim(strMessage) = "/reset" OR Left(strMessage, 5) = "/kick" OR Left(strMessage, 3) = "/pm" OR Left(strMessage, 4) = "/ban" OR Trim(strMessage) = "/logout" Then

	'Only read in the username when its needed to prevent errors
	If Trim(strMessage) <> "/reset" AND Trim(strMessage) <> "/logout" Then

		'Read in the username
		strUsername = Mid(strMessage, InStr(1, strMessage, " ", 1) + 1, Len(strMessage))

		'Remove whitespace
		strUsername = Trim(strUsername)

	End If

	'Login the admin
	If Left(strMessage, 6) = "/admin" Then

		'Read in the password
		strPassword = Mid(strMessage, InStr(1, strMessage, " ", 1) + 1, Len(strMessage))

		'Check the login
		Call loginAdmin(strPassword)

	'Reset the chatroom
	ElseIf Trim(strMessage) = "/reset" AND blnAdmin Then

		'Reset the chatroom
		Call resetFGCWebChat()

	'Post Private Message
	ElseIf Left(strMessage, 3) = "/pm" Then
	
	
	
	
	
	
	ElseIf Left(strMessage, 5) = "/kick" AND blnAdmin Then

	If CheckUsername(strUsername) Then

	
	For intArrayPass = 1 To UBound(saryWebChatUsers, 2)
		If saryWebChatUsers(1, intArrayPass) = strUsername Then
		intArrayNo = intArrayPass
		Exit For
		End If
	Next
	If Not intArrayNo="" Then
	If Not saryWebChatUsers(6, intArrayNo)=True Then

		'Logout the user
		Call logoutUser(strUsername)

		'Post a message
		Call postMessage("", strUsername & " Nickli Kullanýcý " & strLoggedInUsername&" Tarafýndan Odadan Atýlmýþtýr.", "", strFGCWebChatRed, "b")
	End If
	End If
	End If

	'Banned User
	ElseIf Left(strMessage, 4) = "/ban" AND blnAdmin Then
	If CheckUsername(strUsername) Then

	
	For intArrayPass = 1 To UBound(saryWebChatUsers, 2)
		If saryWebChatUsers(1, intArrayPass) = strUsername Then
		intArrayNo = intArrayPass
		Exit For
		End If
	Next
	If Not intArrayNo="" Then
	If Not saryWebChatUsers(6, intArrayNo)=True Then
	
	Set nesne = Server.CreateObject("Scripting.FileSystemObject")
	If nesne.FileExists(Server.MapPath("banneduserid.txt"))=False Then
	Set yaz = nesne.CreateTextFile(Server.MapPath("banneduserid.txt"),True)
	yaz.write(strUsername&"|")
	yaz.close
	Set yaz=Nothing
	Else
	Set ag = nesne.OpenTextFile(Server.MapPath("banneduserid.txt"),8,-2)
	ag.Write(strUsername&"|")
	ag.Close
	Set ag=Nothing
	End If
	Set nesne=Nothing

		'Logout the user
		Call logoutUser(strUsername)

		'Post a message
		Call postMessage("", strUsername & " Nickli Kullanýcýnýn Odaya Giriþi " & strLoggedInUsername&" Tarafýndan Yasaklanmýþtýr.", "", strFGCWebChatRed, "b")
	End If
	End If
	End If
	'Logout
	ElseIf Trim(strMessage) = "/logout" Then

		'Logout the user
		Call logoutUser(strLoggedInUsername)

	End If

'Else post the mesage
Else

	'Select the mode
	Select Case strMode

		'*******************************
		'*** Change the users status ***
		'*******************************
	Case "status"
	Select Case strStatus
	Case strTxtAvailable, strTxtFreeToChat, strTxtBeRightBack, strTxtBusy, strTxtNotAtMyDesk,strTxtOnThePhone
			For intLoop = 1 to UBound(sarySmut)
				strStatus = Replace(strStatus, sarySmut(intLoop, 1), sarySmut(intLoop, 2), 1, -1, 1)
			Next

			'Call the function to change the users status if its not empty
			If NOT isNothing(strStatus) Then Call changeStatus(strStatus)
	End Select
		'**********************
		'*** Post a message ***
		'**********************
	Case "postMessage"

	'Call the function to post the message
	If NOT isNothing(strMessage) Then Call postMessage(strLoggedInUsername, strMessage, strFont, strColor, strFormat)

	End Select

End If
%>