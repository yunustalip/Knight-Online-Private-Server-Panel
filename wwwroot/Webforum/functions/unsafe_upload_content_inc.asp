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



'Dimension variables
Dim saryUnSafeHTMLtags(40)	 'If the following contents is found inside an allowed file type that is not text/HTML based then it could be XSS

'Initalise array values
saryUnSafeHTMLtags(0) = "javascript"
saryUnSafeHTMLtags(1) = "vbscript"
saryUnSafeHTMLtags(2) = "jscript"
saryUnSafeHTMLtags(3) = "object"
saryUnSafeHTMLtags(4) = "applet"
saryUnSafeHTMLtags(5) = "embed"
saryUnSafeHTMLtags(6) = "event"
saryUnSafeHTMLtags(7) = "script"
saryUnSafeHTMLtags(8) = "function"	
saryUnSafeHTMLtags(9) = "cookie"
saryUnSafeHTMLtags(10) = "style"
saryUnSafeHTMLtags(11) = "msgbox"
saryUnSafeHTMLtags(12) = "alert"
saryUnSafeHTMLtags(13) = "create"
saryUnSafeHTMLtags(14) = "hover"
saryUnSafeHTMLtags(15) = "onload"
saryUnSafeHTMLtags(16) = "onclick"
saryUnSafeHTMLtags(17) = "ondblclick"
saryUnSafeHTMLtags(18) = "onkeyup"
saryUnSafeHTMLtags(19) = "onkeydown"
saryUnSafeHTMLtags(20) = "onkeypress"
saryUnSafeHTMLtags(21) = "onkey"
saryUnSafeHTMLtags(22) = "onmouseenter"
saryUnSafeHTMLtags(23) = "onmouseleave"
saryUnSafeHTMLtags(24) = "onmousemove"
saryUnSafeHTMLtags(25) = "onmouseout"
saryUnSafeHTMLtags(26) = "onmouseover"
saryUnSafeHTMLtags(27) = "onrollover"
saryUnSafeHTMLtags(28) = "onmouse"
saryUnSafeHTMLtags(29) = "onchange"
saryUnSafeHTMLtags(30) = "onunloadhave"
saryUnSafeHTMLtags(31) = "onunload"
saryUnSafeHTMLtags(32) = "onsubmit"
saryUnSafeHTMLtags(33) = "onselect"
saryUnSafeHTMLtags(34) = "accesskey"
saryUnSafeHTMLtags(35) = "tabindex"
saryUnSafeHTMLtags(36) = "onfocus"
saryUnSafeHTMLtags(37) = "onblur"
saryUnSafeHTMLtags(38) = "onreset"
saryUnSafeHTMLtags(39) = "mocha"
saryUnSafeHTMLtags(40) = "document"




'If you add more don't forget to increase the number in the Dim statement at the top!
%>