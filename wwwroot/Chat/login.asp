<!--#include file="../_inc/conn.asp"-->
<!--#include file="config.asp" -->
<%dim charids,useridbanned,nesne,ao,veri,xy

'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension Variables
Dim strUsername
Dim intChatroom
Dim strAvatar
Dim strCustomAvatar
Dim strTheme
Dim intLoop
Dim AvatarIs
Dim IntLoops

if Session("login")="ok" Then
'Intialise Variables


'Creat Objects


'Read in requests
strUsername = Trim(Request.Form("nickname"))
strAvatar = Trim(Request.Form("avatar"))
strCustomAvatar = Trim(Request.Form("custom_avatar"))
strTheme = Request.Form("theme")

intChatroom = 1

'Set the theme
Session("ChatChatRoomTheme") = strTheme

'Check for username
If isNothing(strUsername) Then

	'Redirect back to the login page
	Redirect("default.asp?nickname=" & Server.URLEncode(strUsername) & "&MSG=NKNI")

'Else if the username is in use
ElseIf CheckUsername(strUsername) Then

	'Redirect back to the login page
	Redirect("default.asp?nickname=" & Server.URLEncode(strUsername) & "&MSG=NKN")

'Else login the user in the chatroom
Else

	'Format the username
	strUsername = formatUsername(strUsername)
	
	Set charids=Conne.Execute("select * from account_char where straccountid='"&Session("username")&"' and strcharid1='"&strUsername&"' or straccountid='"&Session("username")&"' and strcharid2='"&strUsername&"' or straccountid='"&Session("username")&"' and strcharid3='"&strUsername&"'")
	if not charids.eof Then

	useridbanned=false
	set nesne = Server.CreateObject("Scripting.FileSystemObject")
	set AO = nesne.OpenTextFile(Server.MapPath("banneduserid.txt"),1)
	if not ao.atendofstream Then
	veri=split(ao.readall, "|")
	for xy=0 to ubound(veri)
	if ucase(trim(veri(xy)))=ucase(trim(strUsername)) Then
	useridbanned=true
	Redirect("no_entry.asp")
	Response.End
	exit for
	End If
	next
	End If
	if useridbanned=true Then
	Redirect("no_entry.asp")
	End If

	AvatarIs=0
	
	'Check for custom avatar
	For intLoops = 1 To 30
	If strAvatar = strImagePath & "avatar/avatar" & intLoops & ".gif" Then
	AvatarIs=1
	Exit For
	End If
	Next
	
	If AvatarIs=1 Then

	'Create a username Session
	Session("ChatChatRoomUsername") = strUsername

	'Add the user to the list of users
	Call AddUser(strUsername, intChatroom, strTxtAvailable, strAvatar)
	
	'Redirect to the chatroom
	Redirect("default.asp")
	Else
	Response.Write("<script>alert('Böyle Bir Avatar Seçemezsiniz');location.href='default.asp';</script>")
	Response.End
	End If
	else
	Redirect("default.asp?nickname=" & Server.URLEncode(strUsername) & "&MSG=Yanliskullanici")
	End If
End If
else
Redirect("default.asp?nickname=" & Server.URLEncode(strUsername) & "&MSG=Girisyap")
End If
%>