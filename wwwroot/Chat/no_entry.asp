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
Dim strMode


'Intialise Variables


'Creat Objects


'Read in requests
strMode = Trim(Request.Querystring("M"))


%>
<html>
<head>
<meta name="copyright" content="Copyright (C) 2005 Felix Akinyemi" />
<title><% = strWebsiteName & " - " & strTxtChatroom %></title>


<!--#include file="includes/header.asp"-->
<%

Response.Write(vbCrLf & "<table width=""100%"" height=""100%"">")
Response.Write(vbCrLf & "<tr>")
Response.Write(vbCrLf & "	<td align=""center"">")
Response.Write(vbCrLf & "	<table width=""80%"" height=""80%"" bgcolor=""" & strTableTitleColour2 & """ cellpadding=""5"" cellspacing=""15"">")
Response.Write(vbCrLf & "	<tr>")
Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """ class=""bold"" align=""center"">")
Response.Write(vbCrLf & "		<img src=""" & strImagePath & "no-access.gif"" border=""0"" /><br />" & strTxtUsernameBannedLong & "<br /><br /><br /><br />")

'If the IP Address is banned
If strMode = "IP" Then
	Response.Write(strTxtIPBannedLong)

'If the username is banned
ElseIf strMode = "unameBan" Then
	Response.Write(strTxtUsernameBannedLong)
	Response.Write("<br /><br /><a href=""mailto:" & strWebsiteEmail & """>" & strWebsiteEmail & "</a>")

End If

Response.Write(vbCrLf & "		<br /><br /><br /><input type=""button"" value=""" & strTxtExit & """ class=""bold"" onClick=""window.close();"" />")
Response.Write(vbCrLf & "		</td>")
Response.Write(vbCrLf & "	</tr>")
Response.Write(vbCrLf & "	</table>")
Response.Write(vbCrLf & "	</td>")
Response.Write(vbCrLf & "</tr>")
Response.Write(vbCrLf & "</table>")

%>
<!--#include file="includes/footer.asp"-->