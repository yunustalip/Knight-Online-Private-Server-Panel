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
Dim saryUnSafeHTMLtags(214)	 'If you add more disallowed HTML tags then increase the array size

'Initalise array values
saryUnSafeHTMLtags(0) = "xhtml"
saryUnSafeHTMLtags(1) = "html"
saryUnSafeHTMLtags(2) = "body"
saryUnSafeHTMLtags(3) = "head"
saryUnSafeHTMLtags(4) = "meta"
saryUnSafeHTMLtags(5) = "XSS"
saryUnSafeHTMLtags(6) = "input"
saryUnSafeHTMLtags(7) = "type"
saryUnSafeHTMLtags(8) = "select"
saryUnSafeHTMLtags(9) = "file"
saryUnSafeHTMLtags(10) = "hidden"
saryUnSafeHTMLtags(11) = "checkbox"
saryUnSafeHTMLtags(12) = "password"
saryUnSafeHTMLtags(13) = "checked"
saryUnSafeHTMLtags(14) = "fieldset"
saryUnSafeHTMLtags(15) = "language"
saryUnSafeHTMLtags(16) = "javascript"
saryUnSafeHTMLtags(17) = "vbscript"
saryUnSafeHTMLtags(18) = "jscript"
saryUnSafeHTMLtags(19) = "object"
saryUnSafeHTMLtags(20) = "applet"
saryUnSafeHTMLtags(21) = "embed"
saryUnSafeHTMLtags(22) = "event"
saryUnSafeHTMLtags(23) = "script"
saryUnSafeHTMLtags(24) = "function"	
saryUnSafeHTMLtags(25) = "document"
saryUnSafeHTMLtags(26) = "cookie"
saryUnSafeHTMLtags(27) = "onclick"
saryUnSafeHTMLtags(28) = "ondblclick"
saryUnSafeHTMLtags(29) = "onkeyup"
saryUnSafeHTMLtags(30) = "onkeydown"
saryUnSafeHTMLtags(31) = "onkeypress"
saryUnSafeHTMLtags(32) = "onkey"
saryUnSafeHTMLtags(33) = "onmouseenter"
saryUnSafeHTMLtags(34) = "onmouseleave"
saryUnSafeHTMLtags(35) = "onmousemove"
saryUnSafeHTMLtags(36) = "onmouseout"
saryUnSafeHTMLtags(37) = "onmouseover"
saryUnSafeHTMLtags(38) = "onrollover"
saryUnSafeHTMLtags(39) = "onmouse"
saryUnSafeHTMLtags(40) = "onchange"
saryUnSafeHTMLtags(41) = "onunloadhave"
saryUnSafeHTMLtags(42) = "onunload"
saryUnSafeHTMLtags(43) = "onsubmit"
saryUnSafeHTMLtags(44) = "onselect"
saryUnSafeHTMLtags(45) = "accesskey"
saryUnSafeHTMLtags(46) = "tabindex"
saryUnSafeHTMLtags(47) = "onfocus"
saryUnSafeHTMLtags(48) = "onblur"
saryUnSafeHTMLtags(49) = "onsubmit"
saryUnSafeHTMLtags(50) = "onreset"
saryUnSafeHTMLtags(51) = "form"
saryUnSafeHTMLtags(52) = "iframe"
saryUnSafeHTMLtags(53) = "ilayer"
saryUnSafeHTMLtags(54) = "textarea"
saryUnSafeHTMLtags(55) = "action"
saryUnSafeHTMLtags(56) = "enctype"
saryUnSafeHTMLtags(57) = "layer"
saryUnSafeHTMLtags(58) = "multicol"
saryUnSafeHTMLtags(59) = "frameset"
saryUnSafeHTMLtags(60) = "marquee"
saryUnSafeHTMLtags(61) = "blink"
saryUnSafeHTMLtags(62) = "filter"
saryUnSafeHTMLtags(63) = "overlay"
saryUnSafeHTMLtags(64) = "param"
saryUnSafeHTMLtags(65) = "bgsound"
saryUnSafeHTMLtags(66) = "behavior"
saryUnSafeHTMLtags(67) = "ismap"
saryUnSafeHTMLtags(68) = "sound"
saryUnSafeHTMLtags(69) = "disabled"
saryUnSafeHTMLtags(70) = "ENCTYPE"
saryUnSafeHTMLtags(71) = "!DOCTYPE"
saryUnSafeHTMLtags(72) = "BACKGROUND-COLOR"
saryUnSafeHTMLtags(73) = "base"
saryUnSafeHTMLtags(74) = "position"
saryUnSafeHTMLtags(75) = "absolute"
saryUnSafeHTMLtags(76) = "z-index"
saryUnSafeHTMLtags(77) = "isindex"
saryUnSafeHTMLtags(78) = "xhtml"
saryUnSafeHTMLtags(79) = "xml"
saryUnSafeHTMLtags(80) = "class"
saryUnSafeHTMLtags(81) = "map"
saryUnSafeHTMLtags(82) = "option"
saryUnSafeHTMLtags(83) = "box"
saryUnSafeHTMLtags(84) = "/style" 'Don't want to stripe all 'style' tags as used within span and div area to format text, but if there is a closing </style> tag then it could be a XSS hack
saryUnSafeHTMLtags(85) = "data"
saryUnSafeHTMLtags(86) = "frame"
saryUnSafeHTMLtags(87) = "hspace"
saryUnSafeHTMLtags(88) = "vspace"
saryUnSafeHTMLtags(89) = "css"
saryUnSafeHTMLtags(90) = "float"
saryUnSafeHTMLtags(91) = "SAMP"
saryUnSafeHTMLtags(92) = "link"
saryUnSafeHTMLtags(93) = "alert"	
saryUnSafeHTMLtags(94) = "refresh"
saryUnSafeHTMLtags(95) = "http-equiv"
saryUnSafeHTMLtags(96) = "stylesheet"
saryUnSafeHTMLtags(97) = "blur"
saryUnSafeHTMLtags(98) = "close"
saryUnSafeHTMLtags(99) = "opener"
saryUnSafeHTMLtags(100) = "open"
saryUnSafeHTMLtags(101) = "window"
saryUnSafeHTMLtags(102) = "server"
saryUnSafeHTMLtags(103) = "reset"
saryUnSafeHTMLtags(104) = "RegExp"
saryUnSafeHTMLtags(105) = "protocol"
saryUnSafeHTMLtags(106) = "port"
saryUnSafeHTMLtags(107) = "background"
saryUnSafeHTMLtags(108) = "bgColor"
saryUnSafeHTMLtags(109) = "image"
saryUnSafeHTMLtags(110) = "src"
saryUnSafeHTMLtags(111) = "CharCode"
saryUnSafeHTMLtags(112) = "String"
saryUnSafeHTMLtags(113) = "mocha"
saryUnSafeHTMLtags(114) = "write"
saryUnSafeHTMLtags(115) = "onload"
saryUnSafeHTMLtags(116) = "getElementById"
saryUnSafeHTMLtags(117) = "this"
saryUnSafeHTMLtags(118) = "hover"
saryUnSafeHTMLtags(119) = "image"
saryUnSafeHTMLtags(120) = "visited"
saryUnSafeHTMLtags(121) = "url"
saryUnSafeHTMLtags(122) = "screenposition"
saryUnSafeHTMLtags(123) = "visible"
saryUnSafeHTMLtags(124) = "content"
saryUnSafeHTMLtags(125) = "PLAINTEXT"
saryUnSafeHTMLtags(126) = "button"
saryUnSafeHTMLtags(127) = "radio"
saryUnSafeHTMLtags(128) = "FSCommand"
saryUnSafeHTMLtags(129) = "onabort"
saryUnSafeHTMLtags(130) = "onactivate"
saryUnSafeHTMLtags(131) = "onafter"
saryUnSafeHTMLtags(132) = "charset"
saryUnSafeHTMLtags(133) = "onbegin"
saryUnSafeHTMLtags(134) = "onbounce"
saryUnSafeHTMLtags(135) = "onCellChange"
saryUnSafeHTMLtags(136) = "onContextMenu"
saryUnSafeHTMLtags(137) = "onControlSelect"
saryUnSafeHTMLtags(138) = "onCopy"
saryUnSafeHTMLtags(139) = "onCut"
saryUnSafeHTMLtags(140) = "onDataAvailable"
saryUnSafeHTMLtags(141) = "onDataSetChanged"
saryUnSafeHTMLtags(142) = "onBeforeCut"
saryUnSafeHTMLtags(143) = "onDataSetComplete"
saryUnSafeHTMLtags(144) = "onDeactivate"
saryUnSafeHTMLtags(145) = "onDrag"
saryUnSafeHTMLtags(146) = "onDrop"
saryUnSafeHTMLtags(147) = "onEnd"
saryUnSafeHTMLtags(148) = "onError"
saryUnSafeHTMLtags(149) = "onFilterChange"
saryUnSafeHTMLtags(150) = "onFinish"
saryUnSafeHTMLtags(151) = "onFocus"
saryUnSafeHTMLtags(152) = "onHelp"
saryUnSafeHTMLtags(153) = "onLayoutComplete"
saryUnSafeHTMLtags(154) = "onLoseCapture"
saryUnSafeHTMLtags(155) = "onBefore"
saryUnSafeHTMLtags(156) = "onMediaComplete"
saryUnSafeHTMLtags(157) = "onMediaError"
saryUnSafeHTMLtags(158) = "onMove"
saryUnSafeHTMLtags(159) = "onOutOfSync"
saryUnSafeHTMLtags(160) = "onPaste"
saryUnSafeHTMLtags(161) = "onPause"
saryUnSafeHTMLtags(162) = "onProgress"
saryUnSafeHTMLtags(163) = "onPropertyChange"
saryUnSafeHTMLtags(164) = "onReadyStateChange"
saryUnSafeHTMLtags(165) = "onRepeat"
saryUnSafeHTMLtags(166) = "onReset"
saryUnSafeHTMLtags(167) = "onResize"
saryUnSafeHTMLtags(168) = "onResume"
saryUnSafeHTMLtags(169) = "onReverse"
saryUnSafeHTMLtags(170) = "onRowsEnter"
saryUnSafeHTMLtags(171) = "onRowExit"
saryUnSafeHTMLtags(172) = "onRowDelete"
saryUnSafeHTMLtags(173) = "onRowInserted"
saryUnSafeHTMLtags(174) = "onRow"
saryUnSafeHTMLtags(175) = "onScroll"
saryUnSafeHTMLtags(176) = "onSeek"
saryUnSafeHTMLtags(177) = "onSelect"
saryUnSafeHTMLtags(178) = "onSelectionchange"
saryUnSafeHTMLtags(179) = "onSelectStart"
saryUnSafeHTMLtags(180) = "onStart"
saryUnSafeHTMLtags(181) = "onStop"
saryUnSafeHTMLtags(182) = "onSyncRestored"
saryUnSafeHTMLtags(183) = "onTimeError"
saryUnSafeHTMLtags(184) = "onTrackChange"
saryUnSafeHTMLtags(185) = "onURLFlip"
saryUnSafeHTMLtags(186) = "seekSegmentTime"
saryUnSafeHTMLtags(187) = "msgbox"
saryUnSafeHTMLtags(188) = "expression"
saryUnSafeHTMLtags(189) = "endif"
saryUnSafeHTMLtags(190) = "classid"
saryUnSafeHTMLtags(191) = "clsid"
saryUnSafeHTMLtags(192) = "eval"
saryUnSafeHTMLtags(193) = "namespace"
saryUnSafeHTMLtags(194) = "import"
saryUnSafeHTMLtags(195) = "implementation"
saryUnSafeHTMLtags(196) = "javas"
saryUnSafeHTMLtags(197) = "constructor"
saryUnSafeHTMLtags(198) = "bindings"
saryUnSafeHTMLtags(199) = "binding"
saryUnSafeHTMLtags(200) = "exec"
saryUnSafeHTMLtags(201) = "include"
saryUnSafeHTMLtags(202) = "echo"
saryUnSafeHTMLtags(203) = "unescape"
saryUnSafeHTMLtags(204) = "/*"
saryUnSafeHTMLtags(205) = "*/"
saryUnSafeHTMLtags(206) = "*"
saryUnSafeHTMLtags(207) = "/n"
saryUnSafeHTMLtags(208) = "\"
saryUnSafeHTMLtags(209) = "["
saryUnSafeHTMLtags(210) = "]"
saryUnSafeHTMLtags(211) = "{"
saryUnSafeHTMLtags(212) = "}"
saryUnSafeHTMLtags(213) = "("
saryUnSafeHTMLtags(214) = ")"





'If you add more disallowed HTML tags don't forget to increase the number in the Dim statement at the top!
%>