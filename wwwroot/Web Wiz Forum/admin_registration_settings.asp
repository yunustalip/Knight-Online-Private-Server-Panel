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
blnAvatar = BoolC(Request.Form("avatar"))
blnLongRegForm = BoolC(Request.Form("reg"))
blnModeratorProfileEdit = BoolC(Request.Form("modEdit"))
intMinPasswordLength = IntC(Request.Form("minPass"))
intMinUsernameLength = IntC(Request.Form("minUser"))
blnEnforceComplexPasswords = BoolC(Request.Form("passComplxity"))
blnRealNameReq = BoolC(Request.Form("realName"))
blnLocationReq = BoolC(Request.Form("location"))
blnSignatures = BoolC(Request.Form("signatures"))
blnHomePage = BoolC(Request.Form("homepageURL"))

strCustRegItemName1 = Trim(Request.Form("custRegItemName1"))
blnReqCustRegItemName1 = BoolC(Request.Form("reqCustRegItemName1"))
blnViewCustRegItemName1 = BoolC(Request.Form("viewCustRegItemName1"))
strCustRegItemName2 = Trim(Request.Form("custRegItemName2"))
blnReqCustRegItemName2 = BoolC(Request.Form("reqCustRegItemName2"))
blnViewCustRegItemName2 = BoolC(Request.Form("viewCustRegItemName2"))
strCustRegItemName3 = Trim(Request.Form("custRegItemName3"))
blnReqCustRegItemName3 = BoolC(Request.Form("reqCustRegItemName3"))
blnViewCustRegItemName3 = BoolC(Request.Form("viewCustRegItemName3"))



strRegistrationRules = Request.Form("registrationRules")


'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then	
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	
	Call addConfigurationItem("Avatar", blnAvatar)
	Call addConfigurationItem("Long_reg", blnLongRegForm)
	Call addConfigurationItem("Mod_profile_edit", blnModeratorProfileEdit)
	Call addConfigurationItem("Min_password_length", intMinPasswordLength)
	Call addConfigurationItem("Min_usename_length", intMinUsernameLength)
	Call addConfigurationItem("Password_complexity", blnEnforceComplexPasswords)
	Call addConfigurationItem("Real_name", blnRealNameReq)
	Call addConfigurationItem("Location", blnLocationReq)
	Call addConfigurationItem("Signatures", blnSignatures)
	Call addConfigurationItem("Homepage", blnHomePage)
	Call addConfigurationItem("Registration_Rules", strRegistrationRules)
	
	Call addConfigurationItem("Cust_item_name_1", strCustRegItemName1)
	Call addConfigurationItem("Cust_item_name_req_1", blnReqCustRegItemName1)
	Call addConfigurationItem("Cust_item_name_view_1", blnViewCustRegItemName1)
	Call addConfigurationItem("Cust_item_name_2", strCustRegItemName2)
	Call addConfigurationItem("Cust_item_name_req_2", blnReqCustRegItemName2)
	Call addConfigurationItem("Cust_item_name_view_2", blnViewCustRegItemName2)
	Call addConfigurationItem("Cust_item_name_3", strCustRegItemName3)
	Call addConfigurationItem("Cust_item_name_req_3", blnReqCustRegItemName3)
	Call addConfigurationItem("Cust_item_name_view_3", blnViewCustRegItemName3)
	

		
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnAvatar") = CBool(blnAvatar)
	Application(strAppPrefix & "blnLongRegForm") = CBool(blnLongRegForm)
	Application(strAppPrefix & "blnModeratorProfileEdit") = CBool(blnModeratorProfileEdit)
	Application(strAppPrefix & "intMinPasswordLength") = Cint(intMinPasswordLength)
	Application(strAppPrefix & "intMinUsernameLength") = Cint(intMinUsernameLength)
	Application(strAppPrefix & "blnEnforceComplexPasswords") = CBool(blnEnforceComplexPasswords)
	Application(strAppPrefix & "blnRealNameReq") = CBool(blnRealNameReq)
	Application(strAppPrefix & "blnLocationReq") = CBool(blnLocationReq)
	Application(strAppPrefix & "blnSignatures") = CBool(blnSignatures)
	Application(strAppPrefix & "blnHomePage") = CBool(blnHomePage)
	Application(strAppPrefix & "strRegistrationRules") = strRegistrationRules
	
	Application(strAppPrefix & "strCustRegItemName1") = strCustRegItemName1
	Application(strAppPrefix & "blnReqCustRegItemName1") = CBool(blnReqCustRegItemName1)
	Application(strAppPrefix & "blnViewCustRegItemName1") = CBool(blnViewCustRegItemName1)
	Application(strAppPrefix & "strCustRegItemName2") = strCustRegItemName2
	Application(strAppPrefix & "blnReqCustRegItemName2") = CBool(blnReqCustRegItemName2)
	Application(strAppPrefix & "blnViewCustRegItemName2") = CBool(blnViewCustRegItemName2)
	Application(strAppPrefix & "strCustRegItemName3") = strCustRegItemName3
	Application(strAppPrefix & "blnReqCustRegItemName3") = CBool(blnReqCustRegItemName3)
	Application(strAppPrefix & "blnViewCustRegItemName3") = CBool(blnViewCustRegItemName3)
	
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
	blnAvatar = CBool(getConfigurationItem("Avatar", "bool"))
	blnLongRegForm = CBool(getConfigurationItem("Long_reg", "bool"))
	blnModeratorProfileEdit = CBool(getConfigurationItem("Mod_profile_edit", "bool"))
	intMinPasswordLength = CInt(getConfigurationItem("Min_password_length", "numeric"))
	intMinUsernameLength = CInt(getConfigurationItem("Min_usename_length", "numeric"))
	blnEnforceComplexPasswords = CBool(getConfigurationItem("Password_complexity", "bool"))
	blnRealNameReq = CBool(getConfigurationItem("Real_name", "bool"))
	blnLocationReq = CBool(getConfigurationItem("Location", "bool"))
	blnSignatures = CBool(getConfigurationItem("Signatures", "bool"))
	blnHomePage = CBool(getConfigurationItem("Homepage", "bool"))
	
	strCustRegItemName1 = getConfigurationItem("Cust_item_name_1", "string")
	blnReqCustRegItemName1 = CBool(getConfigurationItem("Cust_item_name_req_1", "bool"))
	blnViewCustRegItemName1 = CBool(getConfigurationItem("Cust_item_name_view_1", "bool"))
	strCustRegItemName2 = getConfigurationItem("Cust_item_name_2", "string")
	blnReqCustRegItemName2 = CBool(getConfigurationItem("Cust_item_name_req_2", "bool"))
	blnViewCustRegItemName2 = CBool(getConfigurationItem("Cust_item_name_view_2", "bool"))
	strCustRegItemName3 = getConfigurationItem("Cust_item_name_3", "string")
	blnReqCustRegItemName3 = CBool(getConfigurationItem("Cust_item_name_req_3", "bool"))
	blnViewCustRegItemName3 = CBool(getConfigurationItem("Cust_item_name_view_3", "bool"))
	
	strRegistrationRules = getConfigurationItem("Registration_Rules", "string")
	
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Kayýt Formu ve Profil Ayarlarý</title>
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
  <h1>Kayýt ve Profil Ayarlarý</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Kontrol Panel Menu</a><br />
    <br />
</div>
<form action="admin_registration_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td colspan="2" class="tableLedger">Kayýt ve Profil Ayarlarý</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Full Kayýt formu:<br />
      <span class="smText">If disabled then new members registering will see a shortened version of the registration form.</span></td>
     <td valign="top" class="tableRow" width="43%">Yes
      <input type="radio" name="reg" value="True" <% If blnLongRegForm = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" value="False" <% If blnLongRegForm = False Then Response.Write "checked" %> name="reg"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Kullanýcý adý uzunluðu (En az):<br />
      <span class="smText">This is minimum length allowed for Members Usernames (max. 20).</span></td>
     <td valign="top" class="tableRow"><select name="minUser" id="minUser"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      <option<% If intMinUsernameLength = 1 Then Response.Write(" selected") %>>1</option>
      <option<% If intMinUsernameLength = 2 Then Response.Write(" selected") %>>2</option>
      <option<% If intMinUsernameLength = 3 Then Response.Write(" selected") %>>3</option>
      <option<% If intMinUsernameLength = 4 Then Response.Write(" selected") %>>4</option>
      <option<% If intMinUsernameLength = 5 Then Response.Write(" selected") %>>5</option>
      <option<% If intMinUsernameLength = 6 Then Response.Write(" selected") %>>6</option>
      <option<% If intMinUsernameLength = 7 Then Response.Write(" selected") %>>7</option>
      <option<% If intMinUsernameLength = 8 Then Response.Write(" selected") %>>8</option>
      <option<% If intMinUsernameLength = 9 Then Response.Write(" selected") %>>9</option>
      <option<% If intMinUsernameLength = 10 Then Response.Write(" selected") %>>10</option>
     </select></td>
    </tr>
    <tr>
     <td class="tableRow">Parola Uzunluðu (En az):<br />
      <span class="smText">This is minimum length allowed for Members Passwords (max. 20).</span></td>
     <td valign="top" class="tableRow"><select name="minPass" id="minPass"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      <option<% If intMinPasswordLength = 1 Then Response.Write(" selected") %>>1</option>
      <option<% If intMinPasswordLength = 2 Then Response.Write(" selected") %>>2</option>
      <option<% If intMinPasswordLength = 3 Then Response.Write(" selected") %>>3</option>
      <option<% If intMinPasswordLength = 4 Then Response.Write(" selected") %>>4</option>
      <option<% If intMinPasswordLength = 5 Then Response.Write(" selected") %>>5</option>
      <option<% If intMinPasswordLength = 6 Then Response.Write(" selected") %>>6</option>
      <option<% If intMinPasswordLength = 7 Then Response.Write(" selected") %>>7</option>
      <option<% If intMinPasswordLength = 8 Then Response.Write(" selected") %>>8</option>
      <option<% If intMinPasswordLength = 9 Then Response.Write(" selected") %>>9</option>
      <option<% If intMinPasswordLength = 10 Then Response.Write(" selected") %>>10</option>
     </select></td>
    </tr>
    <tr>
     <td class="tableRow">Þifre karmaþýlýðý uygulamasý:<br />
       <span class="smText">Bu seçeneði iþaretlerseniz þifre de 1 büyük karakter, 1 küçük karakter, ve 1 numara bulunmasý gerekir (önerilmez).</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="passComplxity" value="True" <% If blnEnforceComplexPasswords = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="passComplxity" value="False" <% If blnEnforceComplexPasswords = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
    <tr>
     <td class="tableRow">Gerçek isim alaný :<br />
       <span class="smText">When enabled members are required to give their Real Name when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="realName" value="True" <% If blnRealNameReq = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="realName" value="False" <% If blnRealNameReq = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
     <tr>
     <td class="tableRow">Konum bilgisi alaný<br />
       <span class="smText">Bu seçenek profile konum bilgisi eklemeye yarar</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="location" value="True" <% If blnLocationReq = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="location" value="False" <% If blnLocationReq = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
     <td class="tableRow">Ýmza :<br />
       <span class="smText">Allow members to create and attach signatures to their Forum Profiles and Posts.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="signatures" value="True" <% If blnSignatures = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="signatures" value="False" <% If blnSignatures = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
     <td class="tableRow">Avatar Resimleri:<br />
       <span class="smText">These are the small images shown next to Members details within the forum system.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="avatar" value="True" <% If blnAvatar = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="avatar" value="False" <% If blnAvatar = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
     <td class="tableRow">Kullanýcý Web Sitesi:<br />
       <span class="smText">Allow members to add a Homepage URL to their websites to be shown within the forum system.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="homepageURL" value="True" <% If blnHomePage = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="homepageURL" value="False" <% If blnHomePage = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
     <td class="tableRow">Modaratörlerin kullanýcý profilini düzenleyebilmesi:<br />
       <span class="smText">When enabled Moderators are able to edit the Forum Profiles of all but the Admin Account. This is useful if an abusive or spamming member needs suspending and the forum admins are not around.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="modEdit" value="True" <% If blnModeratorProfileEdit = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="modEdit" value="False" <% If blnModeratorProfileEdit = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />     </td>
    </tr>
    
    
    <tr>
     <td colspan="2" class="tableLedger">Custom Registration/Profile Item 1</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Item Name:<br />
      <span class="smText">This is the name that you wish displayed for the Custom Registration Item.</span></td>
      <td><input type="text" name="custRegItemName1" id="custRegItemName1" maxlength="25" value="<% = strCustRegItemName1 %>" size="25"<% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
    <td class="tableRow">Required:<br />
       <span class="smText">When enabled members are required to fill in this item when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="reqCustRegItemName1" value="True" <% If blnReqCustRegItemName1 = True Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="reqCustRegItemName1" value="False" <% If blnReqCustRegItemName1 = False Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   <tr>
    <td class="tableRow">Viewed in Member Profile:<br />
       <span class="smText">When enabled all members are able to view this item in member profiles. If disabled only Admins and Moderators can view this item in member profiles.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="viewCustRegItemName1" value="True" <% If blnViewCustRegItemName1 = True Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="viewCustRegItemName1" value="False" <% If blnViewCustRegItemName1 = False Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   
   <tr>
     <td colspan="2" class="tableLedger">Custom Registration/Profile Item 2</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Item Name:<br />
      <span class="smText">This is the name that you wish displayed for the Custom Registration Item.</span></td>
      <td><input type="text" name="custRegItemName2" id="custRegItemName2" maxlength="25" value="<% = strCustRegItemName2 %>" size="25"<% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
    <td class="tableRow">Required:<br />
       <span class="smText">When enabled members are required to fill in this item when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="reqCustRegItemName2" value="True" <% If blnReqCustRegItemName2 = True Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="reqCustRegItemName2" value="False" <% If blnReqCustRegItemName2 = False Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   <tr>
    <td class="tableRow">Viewed in Member Profile:<br />
       <span class="smText">When enabled all members are able to view this item in member profiles. If disabled only Admins and Moderators can view this item in member profiles.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="viewCustRegItemName2" value="True" <% If blnViewCustRegItemName2 = True Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="viewCustRegItemName2" value="False" <% If blnViewCustRegItemName2 = False Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   
   <tr>
     <td colspan="2" class="tableLedger">Custom Registration/Profile Item 3</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Item Name:<br />
      <span class="smText">This is the name that you wish displayed for the Custom Registration Item.</span></td>
      <td><input type="text" name="custRegItemName3" id="custRegItemName3" maxlength="25" value="<% = strCustRegItemName3 %>" size="25"<% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
    <td class="tableRow">Required:<br />
       <span class="smText">When enabled members are required to fill in this item when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="reqCustRegItemName3" value="True" <% If blnReqCustRegItemName3 = True Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="reqCustRegItemName3" value="False" <% If blnReqCustRegItemName3 = False Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   <tr>
    <td class="tableRow">Viewed in Member Profile:<br />
       <span class="smText">When enabled all members are able to view this item in member profiles. If disabled only Admins and Moderators can view this item in member profiles.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="viewCustRegItemName3" value="True" <% If blnViewCustRegItemName3 = True Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="viewCustRegItemName3" value="False" <% If blnViewCustRegItemName3 = False Then Response.Write "checked" %><% If blnDemoMode OR blnACode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>

    <tr>
     <td colspan="2" class="tableLedger">Registration Rules</td>
    </tr>
    <tr>
     <td colspan="2" class="tableRow">
     Enter the Rules new members need to agree to in order to register to become a member (HTML can be used for formatting)
     <br />
     <textarea name="registrationRules" id="registrationRules" rows="15" cols="100"><% = strRegistrationRules %></textarea>
     </td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Kayýt Ayarlarýný Güncelleþtir" />
          <input type="reset" name="Reset" value="Formu Temizle" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
