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




      

'Read in the users details for the forum
blnRTEEditor = BoolC(Request.Form("IEEditor"))
blnFormCAPTCHA = BoolC(Request.Form("CAPTCHA"))
intIncorrectLoginAttempts = IntC(Request.Form("failedLogins"))
blnDisplayTodaysBirthdays = BoolC(Request.Form("showBirthdays"))
blnNewUserCode = BoolC(Request.Form("UserCode"))
blnGuestSessions = BoolC(Request.Form("GuestSID"))
blnEmoticons = BoolC(Request.Form("emoticons"))
blnDisplayMemberList = BoolC(Request.Form("memberList"))
blnModViewIpAddresses = BoolC(Request.Form("modIPs"))
intSearchTimeDefault = IntC(Request.Form("searchDate"))

blnActiveUsers = BoolC(Request.Form("activeUsers"))
blnForumViewing = BoolC(Request.Form("activeUsersViewing"))



'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then	
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	
	Call addConfigurationItem("IE_editor", blnRTEEditor)
	Call addConfigurationItem("Form_CAPTCHA", blnFormCAPTCHA)
	Call addConfigurationItem("Login_attempts", intIncorrectLoginAttempts)
	Call addConfigurationItem("Show_todays_birthdays", blnDisplayTodaysBirthdays)
	Call addConfigurationItem("Tracking_code_update", blnNewUserCode)
	Call addConfigurationItem("Guest_SID", blnGuestSessions)
	Call addConfigurationItem("Emoticons", blnEmoticons)
	Call addConfigurationItem("Show_Member_list", blnDisplayMemberList)
	Call addConfigurationItem("Search_time_default", intSearchTimeDefault)
	
	Call addConfigurationItem("Active_users", blnActiveUsers)
	Call addConfigurationItem("Active_users_viewing", blnForumViewing)
	Call addConfigurationItem("Mod_View_IP", blnModViewIpAddresses)

		
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnRTEEditor") = CBool(blnRTEEditor)
	Application(strAppPrefix & "blnFormCAPTCHA") = CBool(blnFormCAPTCHA)
	Application(strAppPrefix & "intIncorrectLoginAttempts") = Cint(intIncorrectLoginAttempts)
	Application(strAppPrefix & "blnDisplayTodaysBirthdays") = CBool(blnDisplayTodaysBirthdays)
	Application(strAppPrefix & "blnNewUserCode") = CBool(blnNewUserCode)
	Application(strAppPrefix & "blnGuestSessions") = CBool(blnGuestSessions)
	Application(strAppPrefix & "blnEmoticons") = CBool(blnEmoticons)
	Application(strAppPrefix & "blnDisplayMemberList") = CBool(blnDisplayMemberList)
	Application(strAppPrefix & "blnModViewIpAddresses") = CBool(blnModViewIpAddresses)
	Application(strAppPrefix & "intSearchTimeDefault") = Cint(intSearchTimeDefault)
	
	Application(strAppPrefix & "blnActiveUsers") = CBool(blnActiveUsers)
	Application(strAppPrefix & "blnForumViewing") = CBool(blnForumViewing)
	
	Application(strAppPrefix & "blnConfigurationSet") = false
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
	blnRTEEditor =  CBool(getConfigurationItem("IE_editor", "bool"))
	blnFormCAPTCHA = CBool(getConfigurationItem("Form_CAPTCHA", "bool"))
	intIncorrectLoginAttempts = CInt(getConfigurationItem("Login_attempts", "numeric"))
	blnDisplayTodaysBirthdays = CBool(getConfigurationItem("Show_todays_birthdays", "bool"))
	blnNewUserCode = CBool(getConfigurationItem("Tracking_code_update", "bool"))
	blnGuestSessions = CBool(getConfigurationItem("Guest_SID", "bool"))
	blnEmoticons = CBool(getConfigurationItem("Emoticons", "bool"))
	blnDisplayMemberList = CBool(Application(strAppPrefix & "blnDisplayMemberList"))
	blnModViewIpAddresses = CBool(Application(strAppPrefix & "blnModViewIpAddresses"))
	intSearchTimeDefault = CInt(getConfigurationItem("Search_time_default", "numeric"))
	
	blnActiveUsers = CBool(getConfigurationItem("Active_users", "bool"))
	blnForumViewing = CBool(getConfigurationItem("Active_users_viewing", "bool"))
	
	
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Kullanýcý Ayarlarý</title>
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
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
  <h1> Kullanýcý Ayarlarý</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Kontrol Panel Menu</a><br />
    <br />
</div>
<form action="admin_user_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Kullanýcý Ayarlarý</td>
    </tr>
    <tr>
     <td class="tableRow"><a href="http://www.richtexteditor.org" target="_blank">Web Wiz Zengin Yazý Editörü</a> (WYSIWYG Post Editor):<br />
       <span class="smText">This is the type of editor used to create Posts and Private Messages. Requires a Rich Text Enabled web browser, currently Windows IE5+, Firefox 0.6.1+, Netscape 7.1+, Mozilla 1.3+,    
       and Opera 9+, support this feature.<br />
        Members can override this feature by editing their Forum Profile.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="IEEditor" value="True" <% If blnRTEEditor = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="IEEditor" value="False" <% If blnRTEEditor = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Formlarda <a href="http://www.webwizcaptcha.com" target="_blank">Güvenlik Kodu</a> Görüntüle<br />
       <span class="smText">This displays security image on registration, forgotten password, and other forms. Prevents dictionary hacking and remote form submissions.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="CAPTCHA" value="True" <% If blnFormCAPTCHA = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" value="False" <% If blnFormCAPTCHA = False Then Response.Write "checked" %> name="CAPTCHA"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td></tr>
    <tr>
     <td class="tableRow">Baþarýsýz Giriþ Deneme Sayýsý:<br />
      <span class="smText">Belirttiniz sayýdan sonra güvenlik kodu uygulamasý çýkmaktadýr</span></td>
     <td valign="top" class="tableRow">
      <select name="failedLogins" id="failedLogins"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intIncorrectLoginAttempts = 1 Then Response.Write(" selected") %>>1</option>
       <option<% If intIncorrectLoginAttempts = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intIncorrectLoginAttempts = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intIncorrectLoginAttempts = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intIncorrectLoginAttempts = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intIncorrectLoginAttempts = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intIncorrectLoginAttempts = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intIncorrectLoginAttempts = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intIncorrectLoginAttempts = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intIncorrectLoginAttempts = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intIncorrectLoginAttempts = 10 Then Response.Write(" selected") %>>15</option>
       <option<% If intIncorrectLoginAttempts = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intIncorrectLoginAttempts = 50 Then Response.Write(" selected") %>>50</option>
      </select>
     </td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Bugün Doðanlarý Göster<br />
        </td>
      <td width="43%" valign="top" class="tableRow">Evet
       <input type="radio" name="showBirthdays" value="True" <% If blnDisplayTodaysBirthdays = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;Hayýr
        <input type="radio" name="showBirthdays" value="False" <% If blnDisplayTodaysBirthdays = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
     <td class="tableRow">Generate New Tracking Code:<br />
       <span class="smText">Adds extra security by generating a new Tracking Code each time the user logs in, however auto-login on multiple machines will no-longer work.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="userCode" value="True" <% If blnNewUserCode = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="userCode" value="False" <% If blnNewUserCode = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />     </td>
    </tr>
    <tr>
     <td class="tableRow">Guest Sessions:<br />
       <span class="smText">If disabled Guest Users who do not have cookies enabled in their browser can not use some features of the forum, however  it may improve Search Engine Indexing for unknown and rare Search Engine Spiders.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="GuestSID" value="True" <% If blnGuestSessions = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="GuestSID" value="False" <% If blnGuestSessions = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Üye Listesini Göster:<br />
       <span class="smText">When enabled will display the Forum Member List for all Registered Members to view the list of Registered Members.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="memberList" value="True" <% If blnDisplayMemberList = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="memberList" value="False" <% If blnDisplayMemberList = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Moderatörlerin Ip Adreslerini Göster:<br />
       <span class="smText">When enabled will display the IP Addresses of Members and Visitors to Moderators. Forum Admins can always view IP Addresses.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="modIPs" value="True" <% If blnModViewIpAddresses = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="modIPs" value="False" <% If blnModViewIpAddresses = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Smileyleri Aktif Yap<br />
      	 <span class="smText">Enable the use Emoticon Smiley Images within Posts and Private Messages</span></td>
      <td width="43%" valign="top" class="tableRow">Evet
       <input type="radio" name="emoticons" value="True" <% If blnEmoticons = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;Hayýr
      <input type="radio" name="emoticons" value="False" <% If blnEmoticons = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Default Search Time:<br />
      	 <span class="smText">This allows you to select the default search time for forum searches. If you find that searches are taking to long or use to high server resources you should set this to 6 months or below.</span></td>
      <td width="43%" valign="top" class="tableRow">
      <select name="searchDate" id="searchDate"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option value="0"<% If intSearchTimeDefault = 0 Then Response.Write(" selected=""selected""")%>><% = strTxtAnyDate %></option>
       <option value="1"<% If intSearchTimeDefault = 1 Then Response.Write(" selected=""selected""")%>>Last Visit Date</option>
       <option value="2"<% If intSearchTimeDefault = 2 Then Response.Write(" selected=""selected""")%>><% = strTxtYesterday %></option>
       <option value="3"<% If intSearchTimeDefault = 3 Then Response.Write(" selected=""selected""")%>><% = strTxtLastWeek %></option>
       <option value="4"<% If intSearchTimeDefault = 4 Then Response.Write(" selected=""selected""")%>><% = strTxtLastMonth %></option>
       <option value="5"<% If intSearchTimeDefault = 5 Then Response.Write(" selected=""selected""")%>><% = strTxtLastTwoMonths %></option>
       <option value="6"<% If intSearchTimeDefault = 6 Then Response.Write(" selected=""selected""")%>><% = strTxtLastSixMonths %></option>
       <option value="7"<% If intSearchTimeDefault = 7 Then Response.Write(" selected=""selected""")%>><% = strTxtLastYear %>
      </select>
    </tr>
    <tr>
     <td colspan="2" class="tableLedger">Aktif Kullanýcýlar</td>
    </tr>
    <tr>
     <td class="tableRow">Aktif Kullanýcýlar:<br />
       <span class="smText">This displays a list of users presently active within the forums and their location.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="activeUsers" value="True" <% If blnActiveUsers = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="activeUsers" value="False" <% If blnActiveUsers = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Aktif Kullanýcýlarý Forumda Görüntüle:<br />
       <span class="smText">If Active Users (above) is enabled then this option will show how many users are actively viewing each forum.</span></td>
     <td valign="top" class="tableRow">Evet
      <input type="radio" name="activeUsersViewing" value="True" <% If blnForumViewing = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;Hayýr
      <input type="radio" name="activeUsersViewing" value="False" <% If blnForumViewing = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update User Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
