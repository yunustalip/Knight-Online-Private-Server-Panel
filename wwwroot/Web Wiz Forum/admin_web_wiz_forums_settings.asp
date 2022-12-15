<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2011 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  http://www.webwiz.co.uk/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 10E, Dawkins Road Industrial Estate, Poole, Dorset, BH15 4JD, England
'**  http://www.webwiz.co.uk
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************



'*************************** SOFTWARE AND CODE MODIFICATIONS **************************** 
'**
'** MODIFICATION OF THE FREE EDITIONS OF THIS SOFTWARE IS A VIOLATION OF THE LICENSE  
'** AGREEMENT AND IS STRICTLY PROHIBITED
'**
'** If you wish to modify any part of this software a license must be purchased
'**
'****************************************************************************************



'Set the response buffer to true
Response.Buffer = True 


'Dimension variables				
Dim strForumName 		'Holds the forum name
Dim strDBversionInfo		'Holds db infomation



'Read in the db version info
If strDatabaseType = "SQLServer" Then

	strDBversionInfo = sqlServerVersion()
	
End If
      

'Read in the users details for the forum
strForumName = Request.Form("forumName")
strWebSiteName = Request.Form("siteName")
strWebsiteURL = Request.Form("siteURL")
strForumPath = Request.Form("forumPath")
strPageEncoding = Request.Form("pageEncoding")
strTextDirection = Request.Form("direction")
strCookiePrefix = Request.Form("cookiePrefix")
strCookiePath =  Request.Form("cookiePath")
blnDatabaseHeldSessions =  BoolC(Request.Form("session"))
blnShowProcessTime = BoolC(Request.Form("processTime"))
blnDetailedErrorReporting = BoolC(Request.Form("detailedErrors"))
strCookieDomain = Request.Form("cookieDomain")
blnShowLatestPosts = BoolC(Request.Form("showLatestPosts"))
blnHttpXmlApi = BoolC(Request.Form("HttpXmlApi"))
blnDisplayForumStats = BoolC(Request.Form("ForumStats"))
strForumsMessage = Request.Form("forumsMessage")



'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
		
	Call addConfigurationItem("forum_name", strForumName)
	Call addConfigurationItem("website_path", strWebsiteURL)
	Call addConfigurationItem("website_name", strWebSiteName)
	Call addConfigurationItem("forum_path", strForumPath)
	Call addConfigurationItem("Page_encoding", strPageEncoding)
	Call addConfigurationItem("Text_direction", strTextDirection)
	Call addConfigurationItem("Cookie_prefix", strCookiePrefix)
	Call addConfigurationItem("Cookie_path", strCookiePath)
	Call addConfigurationItem("Session_db", blnDatabaseHeldSessions)
	Call addConfigurationItem("Process_time", blnShowProcessTime)
	Call addConfigurationItem("Detailed_error_reporting", blnDetailedErrorReporting)
	Call addConfigurationItem("Cookie_domain", strCookieDomain)
	Call addConfigurationItem("Show_latest_posts", blnShowLatestPosts)
	Call addConfigurationItem("HTTP_XML_API", blnHttpXmlApi)
	Call addConfigurationItem("Show_Forum_Stats", blnDisplayForumStats)
	Call addConfigurationItem("Forums_message", strForumsMessage)
	
		
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "strMainForumName") = strForumName
	Application(strAppPrefix & "strWebsiteURL") = strWebsiteURL
	Application(strAppPrefix & "strWebsiteName") = strWebSiteName
	Application(strAppPrefix & "strForumPath") = strForumPath
	Application(strAppPrefix & "strPageEncoding") = strPageEncoding
	Application(strAppPrefix & "strTextDirection") = strTextDirection
	Application(strAppPrefix & "strCookiePrefix") = strCookiePrefix
	Application(strAppPrefix & "strCookiePath") = strCookiePath
	Application(strAppPrefix & "blnShowProcessTime") = CBool(blnShowProcessTime)
	Application(strAppPrefix & "blnDatabaseHeldSessions") = CBool(blnDatabaseHeldSessions)
	Application(strAppPrefix & "blnDetailedErrorReporting") = CBool(blnDetailedErrorReporting)
	Application(strAppPrefix & "strCookieDomain") = strCookieDomain
	Application(strAppPrefix & "blnShowLatestPosts") = CBool(blnShowLatestPosts)
	Application(strAppPrefix & "blnHttpXmlApi") = CBool(blnHttpXmlApi)
	Application(strAppPrefix & "blnDisplayForumStats") = CBool(blnDisplayForumStats)
	Application(strAppPrefix & "strForumsMessage") = strForumsMessage
	Application.UnLock
End If






'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

	
'Query the database
rsCommon.Open strSQL, adoCon

'Read in the forum from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	
	'Read in the colour info from the database
	strForumName = getConfigurationItem("forum_name", "string")
	strWebsiteURL = getConfigurationItem("website_path", "string")
	strWebSiteName = getConfigurationItem("website_name", "string")
	strForumPath = getConfigurationItem("forum_path", "string")
	strPageEncoding = getConfigurationItem("Page_encoding", "string")
	strTextDirection = getConfigurationItem("Text_direction", "string")
	strCookiePrefix = getConfigurationItem("Cookie_prefix", "string")
	strCookiePath =  getConfigurationItem("Cookie_path", "string")
	blnDatabaseHeldSessions =  CBool(getConfigurationItem("Session_db", "bool"))
	blnModeratorProfileEdit = CBool(getConfigurationItem("Mod_profile_edit", "bool"))
	blnShowProcessTime = CBool(getConfigurationItem("Process_time", "bool"))
	blnDetailedErrorReporting = CBool(getConfigurationItem("Detailed_error_reporting", "bool"))
	strCookieDomain = getConfigurationItem("Cookie_domain", "string")
	blnShowLatestPosts = CBool(getConfigurationItem("Show_latest_posts", "bool"))
	blnHttpXmlApi = CBool(getConfigurationItem("HTTP_XML_API", "bool"))
	blnDisplayForumStats = CBool(getConfigurationItem("Show_Forum_Stats", "bool"))
	strForumsMessage = getConfigurationItem("Forums_message", "string")
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Genel Forum Ayarlarý</title>
<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2011 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript" type="text/javascript">
<!-- Hide from older browsers...

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Check for a name of the forum
	if (document.frmConfiguration.forumName.value==""){
		alert("Please enter the name of your Forum");
		document.frmConfiguration.forumName.focus();
		return false;
	}
	
	//Check for a site name
	if (document.frmConfiguration.siteName.value==""){
		alert("Please enter the name of your Website");
		document.frmConfiguration.siteName.focus();
		return false;
	}
	
	//Check for a URL to homepage
	if (document.frmConfiguration.siteURL.value==""){
		alert("Please enter the URL to your websites homepage");
		document.frmConfiguration.siteURL.focus();
		return false;
	}
	
	//Check for a path to the forum
	if (document.frmConfiguration.forumPath.value==""){
		alert("Please enter the URL to your Forum");
		document.frmConfiguration.forumPath.focus();
		return false;
	}
	
	//Check for a path to the forum
	if (document.frmConfiguration.cookiePrefix.value==""){
		alert("Please enter a Cookie Prefix");
		document.frmConfiguration.cookiePrefix.focus();
		return false;
	}
	
	return true;
}
// -->
</script><!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
  <h1>Genel Forum Ayarlarý</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Kontrol Panel Menu</a><br />
    <br />
    <br />
</div>
<form action="admin_web_wiz_forums_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Forum Ayarlarý</td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Forum Ýsmi*<br />
      <span class="smText">This is the name of your forum. eg. My  Forum </span></td>
      <td width="43%" valign="top" class="tableRow"><input name="forumName" type="text" id="forumName" value="<% = strForumName %>" size="30" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Forum Web Adresi*<br />
       <span class="smText">The web URL to your forum including your domain name and any folder the forum may be in. This is the address you would 
        type into the address bar on your browser to get to the forum.<br />
        eg. http://www.mywebsite.com/forum </span></td>
     <td valign="top" class="tableRow"><input type="text" name="forumPath" maxlength="70" value="<% = strForumPath %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />     </td>
    </tr>
    <tr>
     <td class="tableRow">Sayfa Kodlamasý<br />
      <span class="smText">Unicode UTF-8 supports most languages. If you change the page encoding then some characters maybe rendered incorrectly in browsers and show as (?). This issue can be fixed by resubmitting the data under the new page encoding.</span></td>
     <td valign="top" class="tableRow"><label>
      <select name="pageEncoding" id="pageEncoding"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option value="utf-8"<% If strPageEncoding = "" OR strPageEncoding = "utf-8" Then Response.Write(" selected") %>>Unicode UTF-8</option>
       <option value="iso-8859-1"<% If strPageEncoding = "iso-8859-1" Then Response.Write(" selected") %>>Western European iso-8859-1</option>
       <option value="iso-8859-6"<% If strPageEncoding = "iso-8859-6" Then Response.Write(" selected") %>>Arabic iso-8859-6</option>
       <option value="windows-1256"<% If strPageEncoding = "windows-1256" Then Response.Write(" selected") %>>Arabic windows-1256</option>
       <option value="windows-1257"<% If strPageEncoding = "windows-1257" Then Response.Write(" selected") %>>Baltic windows-1257</option>
       <option value="ibm852"<% If strPageEncoding = "ibm852" Then Response.Write(" selected") %>>Central European DOS ibm852</option>
       <option value="iso-8859-2"<% If strPageEncoding = "iso-8859-2" Then Response.Write(" selected") %>>Central European iso-8859-2</option>
       <option value="windows-1250"<% If strPageEncoding = "windows-1250" Then Response.Write(" selected") %>>Central European windows-1250</option>
       <option value="gb2312"<% If strPageEncoding = "gb2312" Then Response.Write(" selected") %>>Chinese Simplified gb2312</option>
       <option value="hz-gb-2312"<% If strPageEncoding = "hz-gb-2312" Then Response.Write(" selected") %>>Chinese Simplified hz-gb-2312</option>
       <option value="big5"<% If strPageEncoding = "big5" Then Response.Write(" selected") %>>Chinese Traditional big5</option>
       <option value="iso-8859-5"<% If strPageEncoding = "iso-8859-5" Then Response.Write(" selected") %>>Cyrillic iso-8859-5</option>
       <option value="koi8-r"<% If strPageEncoding = "koi8-r" Then Response.Write(" selected") %>>Cyrillic koi8-r</option>
       <option value="koi8-ru"<% If strPageEncoding = "koi8-ru" Then Response.Write(" selected") %>>Cyrillic koi8-ru</option>
       <option value="windows-1251"<% If strPageEncoding = "windows-1251" Then Response.Write(" selected") %>>Cyrillic windows-1251</option>
       <option value="iso-8859-7"<% If strPageEncoding = "iso-8859-7" Then Response.Write(" selected") %>>Greek iso-8859-7</option>
       <option value="windows-1253"<% If strPageEncoding = "windows-1253" Then Response.Write(" selected") %>>Greek windows-1253</option>
       <option value="iso-8859-8-i"<% If strPageEncoding = "iso-8859-8-i" Then Response.Write(" selected") %>>Hebrew ISO-Logical iso-8859-8-i</option>
       <option value="windows-1255"<% If strPageEncoding = "windows-1255" Then Response.Write(" selected") %>>Hebrew windows-1255</option>
       <option value="euc-jp"<% If strPageEncoding = "euc-jp" Then Response.Write(" selected") %>>Japanese EUC euc-jp</option>
       <option value="shift-jis"<% If strPageEncoding = "shift-jis" Then Response.Write(" selected") %>>Japanese Shift-JIS shift-jis</option>
       <option value="euc-kr"<% If strPageEncoding = "euc-kr" Then Response.Write(" selected") %>>Korean euc-kr</option>
       <option value="windows-874"<% If strPageEncoding = "windows-874" Then Response.Write(" selected") %>>Thai windows-874</option>
       <option value="iso-8859-9"<% If strPageEncoding = "iso-8859-9" Then Response.Write(" selected") %>>Turkish iso-8859-9</option>
       <option value="windows-1258"<% If strPageEncoding = "windows-1258" Then Response.Write(" selected") %>>Vietnamese windows-1258</option>
      </select>
     </label></td>
    </tr>
    <tr>
     <td class="tableRow">Yazým Yönü<br />
      <span class="smText">This is the direction your language is written in.</span></td>
     <td valign="top" class="tableRow"><label>
      <select name="direction" id="direction"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option value="ltr"<% If strTextDirection = "" OR strTextDirection = "ltr" Then Response.Write(" selected") %>>Soldan Saða</option>
       <option value="rtl"<% If strTextDirection = "rtl" Then Response.Write(" selected") %>>Saðdan Sola</option>
      </select>
     </label></td>
    </tr>
    <tr>
      <td class="tableRow">Detaylý Sayfa Hata Raporu:<br />
      <span class="smText">Displays a detailed error message if a server error occurs in Web Wiz Forums. If you have a server error then Web Wiz Support Staff will need this detailed error message in order to diagnose the problem.</span></td>
      <td valign="top" class="tableRow">Evet
       <input type="radio" name="detailedErrors" value="True" <% If blnDetailedErrorReporting = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;Hayýr
        <input type="radio" name="detailedErrors" value="False" <% If blnDetailedErrorReporting = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
     <tr>
      <td class="tableRow">Sayfa yüklenme Zamaný Görüntüleme<br />
      <span class="smText">Display in the page footer the time it has taken for the server to generate the page.</span></td>
      <td valign="top" class="tableRow">Evet
       <input type="radio" name="processTime" value="True" <% If blnShowProcessTime = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;Hayýr
        <input type="radio" name="processTime" value="False" <% If blnShowProcessTime = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
     <tr>
     <td class="tableRow">Session larý Veritabanýnda kayýt altýnda tut<br />
     <span class="smText">Session are used to track your visitors and store  data relating to their visit. By storing the data in the database and not the web servers memory prevents sessions from being dropped. <strong>If you change this option you will need to log back in to this Control Panel.</strong></span></td>
     <td valign="top" class="tableRow">Evet
      	<input type="radio" name="session" id="session" value="True"<% If blnDatabaseHeldSessions = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
	&nbsp;&nbsp;Hayýr
	<input type="radio" name="session" id="session" value="False"<% If blnDatabaseHeldSessions = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td class="tableRow">Forum Ýstatistiklerini Göster<br />
     <span class="smText">Displays Forum Stats of your main Forum Index page.</span></td>
     <td valign="top" class="tableRow">Evet
      	<input type="radio" name="ForumStats" id="ForumStats" value="True"<% If blnDisplayForumStats = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
	&nbsp;&nbsp;Hayýr
	<input type="radio" name="ForumStats" id="ForumStats" value="False"<% If blnDisplayForumStats = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td class="tableRow">Son Mesajlarý Göster<br />
     <span class="smText">Displays the a list of the Latest Posts of your main Forum Index page.</span></td>
     <td valign="top" class="tableRow">Evet
      	<input type="radio" name="showLatestPosts" id="showLatestPosts" value="True"<% If blnShowLatestPosts = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
	&nbsp;&nbsp;Hayýr
	<input type="radio" name="showLatestPosts" id="showLatestPosts" value="False"<% If blnShowLatestPosts = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">HTTP XML API:<br />
     <span class="smText">The <a href="HttpAPI.asp" target="_blank" class="smText">HTTP XML API</a> allows you to connect to Web Wiz Forums from 3rd party applications to send and receive data to the Forum Engine. See <a href="http://www.webwiz.co.uk/web-wiz-forums/kb/xml-http-api.htm" target="_blank" class="smText">Web Wiz Forums XML HTTP API</a> for more details.</span></td>
     <td valign="top" class="tableRow">Evet
      	<input type="radio" name="HttpXmlApi" id="HttpXmlApi" value="True"<% If blnHttpXmlApi = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
	&nbsp;&nbsp;Hayýr
	<input type="radio" name="HttpXmlApi" id="HttpXmlApi" value="False"<% If blnHttpXmlApi = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
      <td colspan="2" class="tableLedger">About Your Website</td>
    </tr>
    <tr>
     <td class="tableRow">Website Ismi*<br />
       <span class="smText">The name of your website.  eg. My Website</span></td>
     <td valign="top" class="tableRow"><input type="text" name="siteName" maxlength="50" value="<% = strWebsiteName %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />     </td>
    </tr>
    <tr>
     <td class="tableRow">Website Adresi*<br />
       <span class="smText">This is the URL to your website's homepage.</span></td>
     <td valign="top" class="tableRow"><input name="siteURL" type="text" id="siteURL" value="<% = strWebsiteURL %>" size="30" maxlength="70"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableLedger">Cookies</td>
    </tr>
    <tr>
     <td class="tableRow">Cookie Name Prefix*<br />
      <span class="smText">This is the prefix for any cookies set by the forum. <strong>If you change this option you will need to log back in to this Control Panel.</strong></span></td>
     <td class="tableRow"><input name="cookiePrefix" type="text" id="cookiePrefix" value="<% = strCookiePrefix %>" size="10" maxlength="10"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td class="tableRow">Cookie Path<br />
       <span class="smText">This is the path set for cookies created by the forum. An incorrect path will mean that cookies will not work.</span></td>
      <td class="tableRow"><input name="cookiePath" type="text" id="cookiePath" value="<% = strCookiePath %>" size="30" maxlength="70"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td class="tableRow">Cookie Domain<br />
       <span class="smText">This is the domain the cookies are created for. An incorrect domain will mean that cookies will not work.<strong>This should be left blank unless frames prevent cookies being set correctly for the domain.</strong></span></td>
      <td class="tableRow"><input name="cookieDomain" type="text" id="cookieDomain" value="<% = strCookieDomain %>" size="30" maxlength="70"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableLedger">Forums Message</td>
    </tr>
    <tr>
     <td colspan="2" class="tableRow">
     Forumun Üst Kýsmýnda görüntülenmesini Ýstediðiniz Mesaj Varsa Yazýnýz.
     <br />
     <textarea name="forumsMessage" id="forumsMessage" rows="5" cols="100"><% = strForumsMessage %></textarea>
     </td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Forum Ayarlarýný Güncelleþtir" />
          <input type="reset" name="Reset" value="Formu Temizle" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
