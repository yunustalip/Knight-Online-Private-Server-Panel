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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True

'Dimension variables
Dim iarySpamID	'Array to hold the ID number for each comment to be deleted

'Check the form ID to prevent XCSRF
Call checkFormID(Request.Form("formID"))

'Run through till all checkedspam filters are deleted
For each iarySpamID in Request.Form("chkSpamdID")
	
	'Initalise the strSQL variable with an SQL statement
	strSQL = "DELETE FROM " & strDbTable & "Spam WHERE (Spam_ID ="  & CInt(iarySpamID) & ");"
		
	'execute sql
	adoCon.Execute(strSQL)	
Next
	 
'Reset server variable
Call closeDatabase()


'Return to the spam filter admin page
Response.Redirect("admin_spam_filter_configure.asp" & strQsSID1)
%>
