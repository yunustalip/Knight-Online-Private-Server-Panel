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

'Log out the user
logoutUser(strLoggedInUsername)

'Remove the admin Session
Session(strFGCAppPrefix & "ChatRoomAdmin") = ""
Session(strFGCAppPrefix & "ChatRoomAdmin") = NULL

'Remove the username Session
Session(strFGCAppPrefix & "ChatRoomUsername") = ""
Session(strFGCAppPrefix & "ChatRoomUsername") = NULL

'Remove the last message ID
Session(strFGCAppPrefix & "LastMessageID") = ""
Session(strFGCAppPrefix & "LastMessageID") = NULL

'Redirect to the chatroom
Redirect("default.asp")

%>