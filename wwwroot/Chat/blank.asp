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


'Intialise Variables


'Creat Objects


'Read in requests



%>
<html>
<head>
<meta name="copyright" content="Copyright (C) 2005 Felix Akinyemi" />
<title><% = strWebsiteName & " - " & strTxtChatroom %></title>


<!--#include file="includes/header.asp"-->
<%

'Code here

%>
<!--#include file="includes/footer.asp"-->