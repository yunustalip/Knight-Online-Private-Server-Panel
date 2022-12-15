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


Call closeDatabase()


'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******	

Dim objXMLHTTP, objXmlDoc
Dim strLiName, strLiEM, strLiURL, strXCode, strDataStream, strFID, strFID2
Dim strLicenseServerMsg
Dim intResponseCode
Dim blnUploadComponent
Dim objUpload



strLicenseServerMsg = ""
blnUploadComponent = False
	
		intResponseCode =200
		
	

		

'Upload component check
Private Sub objectCheck(ByRef strComponentName, ByRef strComponent)

	On Error Resume Next
   
	'ASPupload
	Set objUpload = Server.CreateObject(strComponent)
	
	'If an error the componet is not installed
	If Err.Number <> 0 Then
		
		Response.Write(strComponentName & " - Yüklü Deðil")
		
	'Else the component is installed
	Else
		Response.Write(strComponentName & " - <strong>Yüklü</strong>")
		blnUploadComponent = True
	End If

	'Realease Object
	Set objUpload = Nothing
	
	'Disable error handling
	On Error goto 0

End Sub



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="copyright" content="Copyright (C) 2001-2010 Web Wiz" />
<title>Server Requirements Test</title>

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
<h1>Sunucu Gereksinimleri Testi</h1>
 <a href="admin_menu.asp" target="_self">Admin Kontrol Panel Menu</a><br />
 <br />

 <br />
 <table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Upload Bileþenleri Test </td>
  </tr>
  <tr>
   <td class="tableRow">
 <%

Call objectCheck("File System Object (FSO)", "Scripting.FileSystemObject")
Response.Write("<br /><br />")

Call objectCheck("Persits AspUpload 3.x", "Persits.UploadProgress")
Response.Write("<br />")
Call objectCheck("Persits AspUpload", "Persits.Upload.1")
Response.Write("<br />")
Call objectCheck("Dundas Upload", "Dundas.Upload")
Response.Write("<br />")
Call objectCheck("SoftArtisans FileUp (SA FileUp)", "SoftArtisans.FileUp")
Response.Write("<br />")
Call objectCheck("aspSmartUpload", "aspSmartUpload.SmartUpload")
Response.Write("<br />")
Call objectCheck("AspSimpleUpload", "ASPSimpleUpload.Upload")
   
   
%>
   </td>
  </tr>
 </table>
 <br />
 <table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Resim Boyutlandýrma Bileþeni Test</td>
  </tr>
  <tr>
   <td class="tableRow">
    Resimleri yeniden boyutlandýrmak için Persits AspJPEG bileþeni sunucuda yüklü olmalýdýr.
    <br /><br /><%

Call objectCheck("Persits AspJPEG", "Persits.Jpeg")

%>
   </td>
  </tr>
 </table>
 <br />
 <table border="0" cellpadding="4" cellspacing="1" bordercolor="#000000" class="tableBorder">
  <tr>
   <td align="left" class="tableLedger">Email Bileþeni Test </td>
  </tr>
  <tr>
   <td class="tableRow">E-Mail gödnerebilmek için aþaðýdaki bileþenlerden bir tanesi yüklü olmalýdýr
    <br />
    <br /><%

Call objectCheck("CDOSYS", "CDO.Message")
Response.Write("<br />")
Call objectCheck("CDONTS", "CDONTS.NewMail")
Response.Write("<br />")
Call objectCheck("JMail", "JMail.SMTPMail")
Response.Write("<br />")
Call objectCheck("JMail ver.4+", "Jmail.Message")
Response.Write("<br />")
Call objectCheck("AspEmail", "Persits.MailSender")
Response.Write("<br />")
Call objectCheck("AspMail", "SMTPsvg.Mailer")
   
   
%></td>
  </tr>
 </table>
 <br />
 <br />
 <!-- #include file="includes/admin_footer_inc.asp" --><%

'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>