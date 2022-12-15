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




'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******	

Dim objXMLHTTP, objXmlDoc
Dim intResponseCode
Dim strNewVersionNumber
Dim strReleaseDate
Dim strReleaseAbout
Dim strUpdateServerError
Dim strDataStream
Dim strNewsFeed
Dim strLicenseType
Dim strFID, strFID2


strUpdateServerError = ""

strFID = decodeString(strCodeField)
strFID2 = decodeString(strCodeField2)


		intResponseCode = 200
		
		'See if XML reponse is 200
		If intResponseCode = 200 Then 
			
			'Read in the data
			strNewVersionNumber = "10.01"
			strLicenseType = "Premium Edition"
			strReleaseDate = "1st July 2011"
			strReleaseAbout = "New major release with Chat Room, Topic Rating, Post Thanking, and 150 new admin options!"
			strNewsFeed = "Web Wiz Forums version 10 is now released. Download Now! <br>Keep an eye on our Facebook and Twitter pages for more information. "
			
		End If		
			


Call closeDatabase()
	

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="copyright" content="Copyright (C) 2001-2010 Web Wiz" />
<title>Check For Updates</title>

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
<h1>Web Wiz Forum Güncelleþtirmeler</h1>
 <a href="admin_menu.asp" target="_self">Admin Kontrol Paneli</a><br />
 <br />
 <table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Uyarý </td>
  </tr>
  <tr>
   <td class="tableRow">Web Wiz Güncelleþtirmesi yüklemeyiniz crack bozulabilir.
     </td>
  </tr>
 </table>
 <br />
 <br />
 <br />
 <!-- #include file="includes/admin_footer_inc.asp" --><%

'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>