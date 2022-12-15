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
Dim intBadWord		'Loop counter for the bad words
Dim saryNewWord(3)	'Holds the word to enter into db
Dim saryReplaceWord(3)	'Holds the word the swear word is to be replaced with

'Check the form ID to prevent XCSRF
Call checkFormID(Request.Form("formID"))

'Loop round three times to get each new bad word
For intBadWord = 1 to 3

	'Read in the words
	saryNewWord(intBadWord) = Request.Form("badWord" & intBadWord)
	saryReplaceWord(intBadWord) = Request.Form("replaceWord" & intBadWord)
	
	'Escape SQL crashing quotes
	saryNewWord(intBadWord) = Replace(saryNewWord(intBadWord), "'", "''", 1, -1, 1)
	saryReplaceWord(intBadWord) = Replace(saryReplaceWord(intBadWord), "'", "''", 1, -1, 1)

	'Check there is a new bad word and a replacement word to add to the database
	If saryNewWord(intBadWord) <> "" AND saryReplaceWord(intBadWord) <> "" Then

		'Initalise the strSQL variable with an SQL statement
		strSQL = "INSERT INTO " & strDbTable & "Smut (Smut, Word_replace) VALUES ('" & saryNewWord(intBadWord) & "', '" & saryReplaceWord(intBadWord) & "');"
			
		'Write the updated date of last post to the database
		adoCon.Execute(strSQL)		
	End If
Next
	 
'Reset server variable
Call closeDatabase()

'Return to the bad word admin page
Response.Redirect("admin_bad_word_filter_configure.asp" & strQsSID1)
%>
