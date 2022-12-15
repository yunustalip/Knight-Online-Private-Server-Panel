<% @ Language=VBScript
<%Option Explicit%>
<%
Response.Charset = "iso-8859-9"%>
<!--#include file="../sunucuayar.asp"-->
<!--#include file="config.asp" -->
<!--#include file="includes/commands_inc.asp"-->
<!--#include file="includes/chatroom_rules_inc.asp"-->
<%
if menuayar("chat")=1 
	
'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension Variables
Dim intLoop


'Intialise Variables


'Creat Objects


'Read in requests

%>
<html>
<head>
<title><% = strWebsiteName & " - " & strTxtChatroom %></title>
<script type="text/javascript" src="../js/jquery.js"></script>
<script language="JavaScript" src="includes/main.js" type="text/javascript"></script>
<script language="JavaScript">
<!--<%

If isNotLoggedIn 

%>
function viewAvatar(sAvatarPath) {
	if (sAvatarPath != "") getObject("avatarPreview").src = sAvatarPath;
}<%

Else

%>
function togEmoticon() {
	var objEmoticons = getObject("fgcWebChatEmoticons");
	if (objEmoticons.style.display == "none") {
		objEmoticons.style.display = "";
		objEmoticons.style.top = "10%";
		objEmoticons.style.left = "30%";
	} else {
		objEmoticons.style.display = "none";
	}
	document.frmFGCWebChat.message.focus();
}

function addEmoticons(sEmotCode) {
	var objMsgBox = document.frmFGCWebChat.message;
	objMsgBox.value += " " + sEmotCode + " ";
	togEmoticon();
	objMsgBox.focus();
	setEmoticonTip("&nbsp;");
}

function emotHover(ele, sTip) {
	ele.style.border = "1px dotted #C8C8C8";
	ele.style.backgroundColor = "#FAFAFA";
	setEmoticonTip(sTip);
}

function emotOut(ele) {
	ele.style.border = "1px solid #FFFFFF";
	ele.style.backgroundColor = "#FFFFFF";
	setEmoticonTip("&nbsp;");
}

function setEmoticonTip(sTip) {
	getObject("EmoticonTooltip").innerHTML = sTip;
}

function togCommands() {
	var objEmoticons = getObject("fgcWebChatCommands");
	if (objEmoticons.style.display == "none") {
		objEmoticons.style.display = "";
		objEmoticons.style.top = "10%";
		objEmoticons.style.left = "28%";
	} else {
		objEmoticons.style.display = "none";
	}
	document.frmFGCWebChat.message.focus();
}

function postCommand(sCommand) {
	var objMsgBox = document.frmFGCWebChat.message;
	objMsgBox.value = sCommand;
	togCommands();
	objMsgBox.focus();
}

function postmsg() {	
	var objStatus = document.frmFGCWebChat.status.value;
	var objMessage = document.frmFGCWebChat.message;
	var objFont = document.frmFGCWebChat.font.value;
	var objColor = document.frmFGCWebChat.color.value;
	var objFormat = document.frmFGCWebChat.format.value;
	var objChatroom = document.frmFGCWebChat.chatroom.value;

	if (!isEmpty(objMessage.value)) {
		sendRequest("postMessage", objStatus, objMessage.value, objFont, objColor, objFormat, objChatroom);
		objMessage.value = "";
	} else {
		alert("<% = strTxtEnterAMessageToPost %>");
	}
	document.frmFGCWebChat.message.focus();
}

function setStatus() {
	
	var boolPostack = false;
	var sNewStatus = document.frmFGCWebChat.status.value;
	var iChatroom = document.frmFGCWebChat.chatroom.value;
	
	if (!isEmpty(sNewStatus)) {
		boolPostack = true;
		if (boolPostack) {
			sendRequest("status", sNewStatus, "", "", "", "", iChatroom);
		}
	}
	document.frmFGCWebChat.message.focus();
}

function setChatRoom() {
	
	var iNewChatroom = document.frmFGCWebChat.chatroom.value;
	
	if (!isEmpty(iNewChatroom)) {
		sendRequest("chatroom", "", "", "", "", "", iNewChatroom);
	}
	document.frmFGCWebChat.message.focus();
}

function checkImageSize(sImage) {
	if (sImage.width > <% = intImageWidth %>) sImage.style.width = <% = intImageWidth %>;
	if (sImage.height > <% = intImageHeight %>) sImage.style.height = <% = intImageHeight %>;
}

function chatusersreload(){
$.ajax({
   url: 'chatusers.asp',
   success: function(ajaxCevap) {
      $('#chatusers').html(ajaxCevap);
   }
});
}
setInterval(chatusersreload,2000)
function sendRequest(sMode, sStatus, sMessage, sFont, sColor, sFormat, iChatroom) {

	if (!isEmpty(sMode)) {				
	
	$('input#mode').val(sMode);
	$('input#status').val(sStatus);
	$('input#message').val(sMessage);
	$('input#font').val(sFont);
	$('input#color').val(sColor);
	$('input#format').val(sFormat);
	$('input#chatroom').val(iChatroom);
function yolla(){
$.ajax({
   type: 'post',
   url: 'postmessage.asp',
   data: $('#msjfrm').serialize() ,
   success: function(ajaxCevap) {
      $('#chatusers').html(ajaxCevap);
   }
});
}
yolla();
	}
}

<%

End If

%>
-->
</script>
<form name="msjfrm" action="postmessage.asp" method="post" id="msjfrm" >
<input type="hidden" name="mode" id="mode" value="" />
<input type="hidden" name="status" id="status" value="" />
<input type="hidden" name="message" id="message" value="" />
<input type="hidden" name="font" id="font" value="" />
<input type="hidden" name="color" id="color" value="" />
<input type="hidden" name="format" id="format" value="" />
<input type="hidden" name="chatroom" id="chatroom" value="" />
</form>
<!--#include file="includes/header.asp"-->
<%
'Check if the user needs to login
If isNotLoggedIn 
	Response.Write(vbCrLf & "<form name=""frmWebChatLogin"" action=""login.asp"" method=""post"" onSubmit=""if(Trim(document.frmWebChatLogin.nickname.value) != ''){document.frmWebChatLogin.Submit.disabled = true;return true;}else{alert('" & strTxtErrorEnterANickname & "');document.frmWebChatLogin.nickname.focus();return false;}"">")
	Response.Write(vbCrLf & "<input type=""hidden"" name=""chatroom"" value=""1"" />")
	Response.Write(vbCrLf & "<table width=""100%"" height=""100%""><tr><td>")
	Response.Write(vbCrLf & "	<table align=""center"" cellpadding=""5"" cellspacing=""1"" bgcolor=""" & strTableTitleColour2 & """>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td colspan=""2"" align=""center"" class=""chatUsersTitle"" height=""30"">" & strTxtEnterANicknameToLogin & "</td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td align=""right"" bgcolor=""" & strTableColour & """>" & strTxtNickname & "*:</td>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """><input type=""text"" name=""nickname"" id=""nickname"" size=""30"" maxlength=""20"" value=""" & Trim(Request.Querystring("nickname")) & """ /></td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td align=""right"" bgcolor=""" & strTableColour & """>" & strTxtUsersInChatRoom & ":</td>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """>")

	'Check if any users are in the chat room
	If UBound(saryWebChatUsers, 2) > 0 

		'Get the online users
		For intArrayPass = 1 To UBound(saryWebChatUsers, 2)
			'Display the users in chatroom
			Response.Write(saryWebChatUsers(1, intArrayPass))
			'Check if to add a ,
			If UBound(saryWebChatUsers, 2) > 1 AND intArrayPass <> UBound(saryWebChatUsers, 2)  Response.Write(",&nbsp;")
		Next

	'Else there are no users in the chatroom
	Else
		Response.Write(strTxtThereAreNoUsersInChatRoom)
	End If

	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td align=""right"" bgcolor=""" & strTableColour & """ valign=""top"">" & strTxtSelectAvatar & ":<br /><img src=""" & strImagePath & "avatar/avatar1.gif"" id=""avatarPreview"" /></td>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """>")
	Response.Write(vbCrLf & "		  <select name=""avatar"" size=""5"" style=""width: 100%"" class=""smText"" onChange=""viewAvatar(this.value);"">")

	'List the avatars
	For intLoop = 1 To 30
		Response.Write(vbCrLf & "		    <option value=""" & strImagePath & "avatar/avatar" & intLoop & ".gif""") : If intLoop = 1  Response.Write(" selected") End If : Response.Write(">" & strTxtAvatar & " " & intLoop & "</option>")
	Next

	Response.Write(vbCrLf & "		  </select>")
	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td align=""right"" bgcolor=""" & strTableColour & """>" & strTxtSelectTheme & ":</td>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """>")
	Response.Write(vbCrLf & "		   <select name=""theme"" class=""smText"">")
	Response.Write(vbCrLf & "		     <option value=""red""") : If strChatRoomTheme = "red"   Response.Write(" selected") End If : Response.Write(">" & strTxtRed & " (" & strTxtDefault & ")</option>")
	Response.Write(vbCrLf & "		     <option value=""blue""") : If strChatRoomTheme = "blue" OR isNothing(Session(strFGCAppPrefix & "ChatRoomTheme"))  Response.Write(" selected") End If : Response.Write(">" & strTxtBlue & "</option>")
	Response.Write(vbCrLf & "		     <option value=""green""") : If strChatRoomTheme = "green"  Response.Write(" selected") End If : Response.Write(">" & strTxtGreen & "</option>")
	Response.Write(vbCrLf & "		     <option value=""orange""") : If strChatRoomTheme = "orange"  Response.Write(" selected") End If : Response.Write(">" & strTxtOrange & "</option>")
	Response.Write(vbCrLf & "		     <option value=""black""") : If strChatRoomTheme = "black"  Response.Write(" selected") End If : Response.Write(">" & strTxtBlack & "</option>")
	Response.Write(vbCrLf & "		   </select>")
	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """></td>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """>")
	Response.Write(vbCrLf & "		<span class=""bold"">" & strTxtChatRoomRules & "</span>")
	Response.Write(vbCrLf & "	<ol>")

	'Loop through the chatroom rules
	For intLoop = 1 To UBound(saryChatroomRules)
		Response.Write(vbCrLf & "		<li><a href=""javascript:;"" onClick=""alert('" & saryChatroomRules(intLoop, 2) & "');"">" & saryChatroomRules(intLoop, 1) & "</a></li>")
	Next

	Response.Write(vbCrLf & "    </ol>")
	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "	</tr>")	
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """></td>")
	Response.Write(vbCrLf & "		<td bgcolor=""" & strTableColour & """><input type=""submit"" name=""Submit"" value=""" & strTxtIAgreeLogin & """ />&nbsp;<input type=""button"" value=""" & strTxtExit & """ onClick=""window.close();"" /></td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	</table>")
	Response.Write(vbCrLf & "</td></tr></table>")
	Response.Write(vbCrLf & "</form>")
	Response.Write(vbCrLf & "<script>document.frmWebChatLogin.nickname.focus();</script>")

	'Popup messages
	If Trim(Request.Querystring("MSG")) = "NKN" 
		Response.Write(vbCrLf & "<script>alert('" & strTxtErrorNicknameInUse & "');</script>")
	ElseIf Trim(Request.Querystring("MSG")) = "NKNI" 
		Response.Write(vbCrLf & "<script>alert('" & strTxtErrorEnterANickname & "');</script>")
	End If

'Else the user is logged in
Else
if Session("login")="ok" and Session("yetki")="1" 
		Session("ChatChatRoomAdmin") = "1"

		'Set the user to admin
		saryWebChatUsers(6, intArrayPos) = True

		'Update the chatroom
		Call updateChatroom()
End If
	Response.Write(vbCrLf & "<form name=""frmFGCWebChat"" onSubmit=""postmsg();return false;"">")
	Response.Write(vbCrLf & "<input type=""hidden"" id=""chatroom"" value=""1"" />")
	Response.Write(vbCrLf & "<table cellpadding=""3"" cellspacing=""0"" width=""100%"" height=""100%"">")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td width=""25%"" valign=""top"">")
	Response.Write(vbCrLf & "		  <table cellpadding=""3"" cellspacing=""2"" width=""100%"" bgcolor=""" & strTableTitleColour2 & """>")
	Response.Write(vbCrLf & "		  <tr>")
	Response.Write(vbCrLf & "		  	<td class=""chatUsersTitle"" align=""center"" height=""25""><img src=""" & strImagePath & "users.gif"" align=""middle"" />&nbsp;" & strTxtUsersOnline & "</td>")
	Response.Write(vbCrLf & "		  </tr>")
	Response.Write(vbCrLf & "		  <tr>")
	Response.Write(vbCrLf & "		  	<td bgcolor=""" & strTableColour & """><div width=""100%"" height=""370"" name=""chatUsersFrame"" id=""chatusers"">")
	Server.Execute("chatusers.asp")
	Response.Write("</div></td>")
	Response.Write(vbCrLf & "		  </tr>")
	Response.Write(vbCrLf & "		  </table>")
	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "		<td width=""75%"" valign=""top"">")
	Response.Write(vbCrLf & "		  <table cellpadding=""3"" cellspacing=""2"" width=""100%"" bgcolor=""" & strTableTitleColour2 & """>")
	Response.Write(vbCrLf & "		  <tr>")

	'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	If blnCode = True 
		Response.Write(vbCrLf & "		  	<td class=""chatUsersTitle"" height=""25""><a href=""mailto:delikanli-1903@hotmail.com"" target=""_blank"" class=""chatUsersTitle"">Powered by Asi KartaL " & strVersion & "</a></td><td align=""right""><a href=""logout.asp"" onClick=""return confirm('" & strTxtConfirmSignOut & "');""><img src=""" & strImagePath & "sign_out.gif"" border=""0"" align=""absmiddle"" /></a></td>")
	Else
		Response.Write(vbCrLf & "		  	<td class=""chatUsersTitle"" height=""25"">" & strWebsiteName & " " & strVersion & "</a></td></td><td align=""right""><a href=""logout.asp"" onClick=""return confirm('" & strTxtConfirmSignOut & "');""><img src=""" & strImagePath & "sign_out.gif"" border=""0"" align=""absmiddle"" /></a></td>")
	End If
	
	Response.Write(vbCrLf & "		  </tr>")
	Response.Write(vbCrLf & "		  <tr>")
	Response.Write(vbCrLf & "		  	<td class=""chatUsersTitle"" bgcolor=""" & strTableColour & """ colspan=""2""><div class=""chatBody"" id=""chatBody"" style=""overflow:auto;height: 370;width: 100%;color: #000000;font-weight: normal;""></div></td>")
	Response.Write(vbCrLf & "		  </tr>")
	Response.Write(vbCrLf & "		  </table>")
	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "	<tr>")
	Response.Write(vbCrLf & "		<td colspan=""2"" valign=""top"">")
	Response.Write(vbCrLf & "		  <table cellpadding=""3"" cellspacing=""2"" width=""100%"" bgcolor=""" & strTableTitleColour2 & """>")
	Response.Write(vbCrLf & "		  <tr>")
	Response.Write(vbCrLf & "		  	<td class=""chatUsersTitle"" height=""25"">" & strTxtMessage & ":</td><td class=""chatUsersTitle"" height=""25"" align=""right"">")
	Response.Write(vbCrLf & strTxtStatus & ":&nbsp;<select id=""status"" onChange=""setStatus();"" class=""smText"">")
	Response.Write(vbCrLf & "	<option value=""" & strTxtAvailable & """") : If strLoggedInUserStats = strTxtAvailable  Response.Write(" selected") End If : Response.Write(">" & strTxtAvailable & "</option>")
	Response.Write(vbCrLf & "	<option value=""" & strTxtFreeToChat & """") : If strLoggedInUserStats = strTxtFreeToChat  Response.Write(" selected") End If : Response.Write(">" & strTxtFreeToChat & "</option>")
	Response.Write(vbCrLf & "	<option value=""" & strTxtBeRightBack & """") : If strLoggedInUserStats = strTxtBeRightBack  Response.Write(" selected") End If : Response.Write(">" & strTxtBeRightBack & "</option>")
	Response.Write(vbCrLf & "	<option value=""" & strTxtBusy & """") : If strLoggedInUserStats = strTxtBusy  Response.Write(" selected") End If : Response.Write(">" & strTxtBusy & "</option>")
	Response.Write(vbCrLf & "	<option value=""" & strTxtNotAtMyDesk & """") : If strLoggedInUserStats = strTxtNotAtMyDesk  Response.Write(" selected") End If : Response.Write(">" & strTxtNotAtMyDesk & "</option>")
	Response.Write(vbCrLf & "	<option value=""" & strTxtOnThePhone & """") : If strLoggedInUserStats = strTxtOnThePhone  Response.Write(" selected") End If : Response.Write(">" & strTxtOnThePhone & "</option>")
	Response.Write(vbCrLf & "</select>")

	If blnPlayNewMessageSound 
		Response.Write(vbCrLf & strTxtSound & ":&nbsp;<select id=""sound"" class=""smText"" onChange=""document.frmFGCWebChat.message.focus();"">")
		Response.Write(vbCrLf & "	<option value=""on"" selected>" & strTxtOn & "</option>")
		Response.Write(vbCrLf & "	<option value=""off"">" & strTxtOff & "</option>")
		Response.Write(vbCrLf & "</select>")
	End If

	Response.Write("</td>")
	Response.Write(vbCrLf & "		  </tr>")
	Response.Write(vbCrLf & "		  <tr>")
	Response.Write(vbCrLf & "		  	<td class=""chatUsersTitle"" bgcolor=""" & strTableColour & """ valign=""top"" colspan=""2"">")
	Response.Write(vbCrLf & "<table width=""100%"" height=""100%"" cellpadding=""1"" cellspacing=""2"" bgcolor=""" & strTableColour & """>")
	Response.Write(vbCrLf & "<tr height=""40"">")
	Response.Write(vbCrLf & "	<td><input type=""text"" id=""message"" name=""message"" size=""50"" autocomplete=""off"" />&nbsp;<input type=""submit"" value=""" & strTxtSend & """ name=""Submit"" id=""Submit"" class=""handCursor"" /></td>")
	Response.Write(vbCrLf & "	<td>&nbsp;</td>")
	Response.Write(vbCrLf & "	<td width=""1%""><img src=""" & strImagePath & "commands.gif"" border=""0"" class=""handCursor"" onClick=""togCommands();"" title=""" & strTxtChatRoomCommands & """ /></td>")
	Response.Write(vbCrLf & "	<td width=""1%""><img src=""" & strImagePath & "emoticons.gif"" border=""0"" class=""handCursor"" onClick=""togEmoticon();"" title=""" & strTxtEmoticons & """ /></td>")
	Response.Write(vbCrLf & "	<td width=""1%"">")
	Response.Write(vbCrLf & "<select id=""font"" class=""smText"" onChange=""document.frmFGCWebChat.message.focus();"">")
	Response.Write(vbCrLf & "	<option value=""Arial"">Arial</option>")
	Response.Write(vbCrLf & "	<option value=""Book Antiqua"">Book Antiqua</option>")
	Response.Write(vbCrLf & "	<option value=""Bookman Old Style"">Bookman Old Style</option>")
	Response.Write(vbCrLf & "	<option value=""Broadway"">Broadway</option>")
	Response.Write(vbCrLf & "	<option value=""Century Gothic"">Century Gothic</option>")
	Response.Write(vbCrLf & "	<option value=""Comic Sans MS"">Comic Sans MS</option>")
	Response.Write(vbCrLf & "	<option value=""Courier"">Courier</option>")
	Response.Write(vbCrLf & "	<option value=""Garamond"">Garamond</option>")
	Response.Write(vbCrLf & "	<option value=""Gill Sans MT"">Gill Sans MT</option>")
	Response.Write(vbCrLf & "	<option value=""Haettenschweiler"">Haettenschweiler</option>")
	Response.Write(vbCrLf & "	<option value=""Helvetica"">Helvetica</option>")
	Response.Write(vbCrLf & "	<option value=""Impact"">Impact</option>")
	Response.Write(vbCrLf & "	<option value=""Lucida Bright"">Lucida Bright</option>")
	Response.Write(vbCrLf & "	<option value=""Lucida Console"">Lucida Console</option>")
	Response.Write(vbCrLf & "	<option value=""Lucida Sans"">Lucida Sans</option>")
	Response.Write(vbCrLf & "	<option value=""Tahoma"">Tahoma</option>")
	Response.Write(vbCrLf & "	<option value=""Times New Roman"">Times New Roman</option>")
	Response.Write(vbCrLf & "	<option value=""Verdana"" selected>Verdana</option>")
	Response.Write(vbCrLf & "</select>")
	Response.Write(vbCrLf & "	</td>")
	Response.Write(vbCrLf & "	<td width=""1%"">")
	Response.Write(vbCrLf & "<select id=""color"" class=""smText"" onChange=""document.frmFGCWebChat.message.focus();"">")
	Response.Write(vbCrLf & "	<option value=""red"" style=""color: red;"">Kýrmýzý</option>")
	Response.Write(vbCrLf & "	<option value=""blue"" style=""color: blue;"">Mavi</option>")
	Response.Write(vbCrLf & "	<option value=""green"" style=""color: green;"">Yeþil</option>")
	Response.Write(vbCrLf & "	<option value=""yellow"" style=""color: yellow;"">Sarý</option>")
	Response.Write(vbCrLf & "	<option value=""orange"" style=""color: orange;"">Turuncu</option>")
	Response.Write(vbCrLf & "	<option value=""brown"" style=""color: brown;"">Kahverengi</option>")
	Response.Write(vbCrLf & "	<option value=""violet"" style=""color: violet;"">Pembe</option>")
	Response.Write(vbCrLf & "	<option value=""black"" style=""color: black;"" selected>Siyah</option>")
	Response.Write(vbCrLf & "</select>")
	Response.Write(vbCrLf & "	</td>")
	Response.Write(vbCrLf & "	<td width=""1%"">")
	Response.Write(vbCrLf & "<select id=""format"" class=""smText"" onChange=""document.frmFGCWebChat.message.focus();"">")
	Response.Write(vbCrLf & "	<option value=""i"">Italik(Yatýk)</option>")
	Response.Write(vbCrLf & "	<option value=""b"">Kalýn</option>")
	Response.Write(vbCrLf & "	<option value=""u"">Alt çizgili</option>")
	Response.Write(vbCrLf & "	<option value="""" selected>Normal</option>")
	Response.Write(vbCrLf & "</select>")
	Response.Write(vbCrLf & "	</td>")
	Response.Write(vbCrLf & "</tr>")
	Response.Write(vbCrLf & "</table>")
	Response.Write(vbCrLf & "		  	</td>")
	Response.Write(vbCrLf & "		  </tr>")
	Response.Write(vbCrLf & "		  </table>")
	Response.Write(vbCrLf & "		</td>")
	Response.Write(vbCrLf & "	</tr>")
	Response.Write(vbCrLf & "</table>")
	Response.Write(vbCrLf & "</form>")

	'***********************
	'***    Emoticons    ***
	'***********************
	Response.Write(vbCrLf & "<div id=""fgcWebChatEmoticons"" class=""Transparency"" style=""position: absolute; display: none;"">")
	Response.Write(vbCrLf & "<table cellpadding=""8"" cellspacing=""1"" bgcolor=""" & strTableBorderColour & """>")
	Response.Write(vbCrLf & "<tr>")
	Response.Write(vbCrLf & "	<td colspan=""9"" bgcolor=""" & strTableBorderColour & """ class=""chatUsersTitle"">" & strTxtEmoticons & "</td>")
	Response.Write(vbCrLf & "	<td bgcolor=""" & strTableColour & """ align=""center"" onClick=""togEmoticon();"" title=""" & strTxtCloseEmoticons & """ class=""handCursor""><img src=""" & strImagePath & "close.gif"" /></td>")
	Response.Write(vbCrLf & "</tr>")
	Response.Write(vbCrLf & "<tr>")

	'Loop through the smlies
	For intLoop = 1 To UBound(saryEmoticons)

		'Print out the HTML for the emoticon
		Response.Write(vbCrLf & "	<td align=""center"" class=""Emot"" onMouseOver=""emotHover(this, '" & Replace(saryEmoticons(intLoop, 1), "'", "\'", 1, -1, 1) & "');"" onMouseOut=""emotOut(this);"" onClick=""addEmoticons('" & formatJSMessage(saryEmoticons(intLoop, 2)) & "');emotOut(this);""><img src=""" & saryEmoticons(intLoop, 3) & """ border=""0"" align=""absmiddle"" /></td>")

		'Break the table to another row
		If (intLoop MOD 10 = 0) AND intLoop <> UBound(saryEmoticons)  Response.Write("</tr>" & vbCrLf & "<tr>")

	Next

	Response.Write(vbCrLf & "</tr>")
	Response.Write(vbCrLf & "<tr>")
	Response.Write(vbCrLf & "	<td colspan=""10"" bgcolor=""" & strTableBorderColour & """ class=""chatUsersTitle"" id=""EmoticonTooltip"" align=""center"">&nbsp;</td>")
	Response.Write(vbCrLf & "</tr>")
	Response.Write(vbCrLf & "</table>")
	Response.Write(vbCrLf & "</div>")

	'**********************
	'***    Commands    ***
	'**********************
	Response.Write(vbCrLf & "<div id=""fgcWebChatCommands"" class=""Transparency"" style=""position: absolute; display: none; overflow: auto; height: 340;"">")
	Response.Write(vbCrLf & "<table cellpadding=""8"" cellspacing=""1"" bgcolor=""" & strTableBorderColour & """>")
	Response.Write(vbCrLf & "<tr>")
	Response.Write(vbCrLf & "	<td colspan=""2"" bgcolor=""" & strTableBorderColour & """ class=""chatUsersTitle"">" & strTxtChatRoomCommands & "</td>")
	Response.Write(vbCrLf & "	<td bgcolor=""" & strTableColour & """ align=""center"" onClick=""togCommands();"" title=""" & strTxtCloseCommands & """ class=""handCursor""><img src=""" & strImagePath & "close.gif"" /></td>")
	Response.Write(vbCrLf & "</tr>")

	'Loop through the smlies
	For intLoop = 1 To UBound(saryCommand)

		'Print out the HTML for the emoticon
		If (saryCommand(intLoop, 3) AND blnAdmin) OR NOT saryCommand(intLoop, 3) 
			Response.Write(vbCrLf & "<tr title=""" & saryCommand(intLoop, 1) & """>")
			Response.Write(vbCrLf & "	<td class=""Commands"">" & saryCommand(intLoop, 1) & "</td><td class=""Commands"">" & saryCommand(intLoop, 2) & "</td><td class=""Commands"" align=""center""><img src=""" & strImagePath & "use_command.gif"" border=""0"" onClick=""postCommand('" & formatJSMessage(saryCommand(intLoop, 2)) & "');"" class=""handCursor"" title=""" & strTxtUseCommand & """ /></td>")
			Response.Write(vbCrLf & "</tr>")
		End If

	Next

	Response.Write(vbCrLf & "</table>")
	Response.Write(vbCrLf & "</div>")
	Response.Write(vbCrLf & "<script>window.focus();document.frmFGCWebChat.message.focus();</script>")

	'If play sound  print the HTML
	If blnPlayNewMessageSound 
		Response.Write(vbCrLf & "<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0"" width=""10"" height=""10"" id=""newMsgSound"" align=""middle"">")
		Response.Write(vbCrLf & "<param name=""allowScriptAccess"" value=""sameDomain"" />")
		Response.Write(vbCrLf & "<param name=""movie"" value=""flash/newMsgSound.swf"" />")
		Response.Write(vbCrLf & "<param name=""quality"" value=""high"" />")
		Response.Write(vbCrLf & "<param name=""bgcolor"" value=""" & strBgColour & """ />")
		Response.Write(vbCrLf & "<embed src=""flash/newMsgSound.swf"" quality=""high"" bgcolor=""" & strBgColour & """ width=""10"" height=""10"" name=""newMsgSound"" align=""middle"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />")
		Response.Write(vbCrLf & "</object>")
	End If

End If

%>
<!--#include file="includes/footer.asp"-->
 <%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if%>