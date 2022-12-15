<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="functions/functions_common.asp" -->
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


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If




'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******	

Dim objXMLHTTP, objXmlDoc
Dim strLiName, strLiEM, strLiURL, strXCode, strDataStream, strFID, strFID2
Dim strErrorMsg
Dim intResponseCode
strErrorMsg = ""
strFID = decodeString(strCodeField)
strFID2 = decodeString(strCodeField2)
strSQL = "SELECT " & strDbTable & "SetupOptions.* From " & strDbTable & "SetupOptions;"
	
With rsCommon
	
	If Request.Form("postBack") Then	
			
		strLiName = Request.Form("liname")
		strLiEM = Request.Form("email")
		strLiURL = LCase(Trim(Request.Form("URL")))
		strXCode = UCase(Trim(Replace(Request.Form("code"), "'", "", 1, -1, 1)))

				intResponseCode = 200

			
						Call addConfigurationItem(strFID, 0)
						Call addConfigurationItem(strFID2, 0)
						Call addConfigurationItem("Install_ID", "AsiBesiktasli")
				
				
		
		.Open strSQL, adoCon
		saryConfiguration = rsCommon.GetRows()
		.Close
		Application.Lock
		Application(strAppPrefix & "strInstallID") = getConfigurationItem("Install_ID", "string")
		Application(strAppPrefix & "blnLCode") = CBool(getConfigurationItem(strFID, "bool"))
		Application(strAppPrefix & "blnACode") = CBool(getConfigurationItem(strFID2, "bool"))
		Application(strAppPrefix & "blnConfigurationSet") = false
		Application.UnLock
		
	End If
	
	.Open strSQL, adoCon
	saryConfiguration = rsCommon.GetRows()
	If NOT .EOF Then
		blnLCode = CBool(getConfigurationItem(strFID, "bool"))
		blnACode = CBool(getConfigurationItem(strFID2, "bool"))
		strInstallID = getConfigurationItem("Install_ID", "string")
	End If
	.Close
End With
Call closeDatabase()



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="copyright" content="Copyright (C) 2001-2010 Web Wiz" />
<title>Web Wiz Forums Premium Edition Upgrade</title>

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

<script  language="JavaScript">

function submitForm() {

        var errorMsg = "";
        var errorMsgLong = "";
   	var formArea = document.getElementById('frmLinkCode');

        if (formArea.URL.value ==''){
                errorMsg += "\n\URL \t\t- Enter the exact regsitration URL";
        }
        if (formArea.code.value ==''){
                errorMsg += "\n\License Key\t- Enter your exact license key";
        }
        if ((errorMsg != "") || (errorMsgLong != "")){
                msg = "_______________________________________________________________\n\n";
                msg += "The form has not been submitted because there are problem(s) with the form.\n";
                msg += "Please correct the problem(s) and re-submit the form.\n";
                msg += "_______________________________________________________________\n\n";
                msg += "The following field(s) need to be corrected: -\n";

                errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
                return false;
        }
       
       return true;
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<h1>Web Wiz Forums Premium Edition Upgrade </h1>
 <a href="admin_menu.asp<% = strQsSID1 %>" target="_self">Return to the the Admin Control Panel Menu</a><br />
 <br /><%
 
If strErrorMsg <> "" Then
	
%>
 <table class="errorTable" cellspacing="1" cellpadding="3">
  <tr>
   <td align="left"><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>Error</strong></td>
  </tr>
  <tr>
   <td align="left"><strong>The Premium Edition License Key has not been applied</strong><br />
    <br />
    <% = strErrorMsg %>
   </td>
  </tr>
 </table>
 <br /><%

End If

If (blnLCode = False OR blnACode = False) AND strInstallID <> "" Then
%>
<table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Tebrikler</td>
  </tr>
  <tr>
   <td class="tableRow"><span class="text">
     <br />
     Cracked By Bug Researchers Group - Asi_Besiktasli</span></td>
  </tr>
 </table>
 <br /><%

End If

%>
 
 <br />
 <table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Web Wiz Forums Premium Edition Server Requirments Check </td>
  </tr>
  <tr>
   <td class="tableRow"><span class="text">Web Wiz Forums Premium Edition requires that you are using a supported web host. Please use the button below to check that the web server you are using and your web host meet the requirements for running the Premium Edition.<br />
    <br />
   </span>
    <form id="frmTestSvr" name="frmTestSvr" method="post" action="admin_server_test.asp<% = strQsSID1 %>">
     <span class="text">
     <input name="testSvr" type="submit" id="testSvr" value="Web Wiz Forums Server Requirements Test" />
     </span>
    </form>    
   </td>
  </tr>
 </table>
 <br />
 <br /><%
 
If intResponseCode <> 200 Then
	
%>
<form method="post" name="frmLinkCode"  id="frmLinkCode" action="admin_license.asp<% = strQsSID1 %>" onSubmit="return submitForm();">
 <table border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
   <td colspan="2" class="tableLedger">Web Wiz Forums Premium Edition License Key </td>
  </tr>
  <tr>
   <td align="right" valign="top" class="tableRow">Registration URL:</td>
   <td valign="top" class="tableRow"><input name="URL" type="text" id="URL" size="35" maxlength="100" value="<% = strLiURL %>" />
     <br />
   <span class="smText">This needs to be the exact URL you registered in your Web Wiz Client Area for the license key being used.</span></td>
  </tr>
  <tr>
   <td align="right" valign="top" class="tableRow">License Key: </td>
   <td class="tableRow"><input type="text" name="code" size="70" maxlength="70" value="<% = strXCode %>" />
     <br />
   <span class="smText">This needs to be exactly the same license key as in your Web Wiz Client Area.</span></td>
  </tr>
  <tr>
   <td width="23%" align="right" class="tableRow"><input type="hidden" name="postBack" id="postBack" value="true" /><input type="hidden" name="svrReponse" id="svrReponse" value="" />   </td>
   <td width="77%" class="tableRow"><input type="submit" name="Submit" id="Submit" value="Submit Form" />
     <input type="reset" name="Reset" value="Reset Form" />   </td>
  </tr>
 </table>
</form><%

End If

%>
 <br />
 <br />
<!-- #include file="includes/admin_footer_inc.asp" --><%


If intResponseCode = 200 Then
	Response.Write("<script language=""JavaScript"">")
	Response.Write("alert('Cracked by Bug Researchers - Asi_Besiktasli for Cyber-Warrior users.\n\nBug Researchers Group')")
	Response.Write("</script>")
End If



'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>