<%dim strTarayici,strKullanilanTarayici
Function BrowserType()
     strTarayici = Request.ServerVariables("HTTP_USER_AGENT")
 
      '// Opera
     If InStr(1, strTarayici, "Opera 1", 1) > 0 Then
          strKullanilanTarayici = "Opera 1"
 
     ElseIf InStr(1, strTarayici, "Opera 2", 1) > 0 Then
          strKullanilanTarayici = "Opera 2"
 
     ElseIf InStr(1, strTarayici, "Opera 3", 1) > 0 Then
          strKullanilanTarayici = "Opera 3"
 
     ElseIf InStr(1, strTarayici, "Opera 4", 1) > 0 Then
          strKullanilanTarayici = "Opera 4"
 
     ElseIf InStr(1, strTarayici, "Opera 5", 1) > 0 Then
          strKullanilanTarayici = "Opera 5"
 
     ElseIf InStr(1, strTarayici, "Opera 6", 1) > 0 Then
          strKullanilanTarayici = "Opera 6"
 
     ElseIf InStr(1, strTarayici, "Opera 7", 1) > 0 Then
          strKullanilanTarayici = "Opera 7"
 
     ElseIf InStr(1, strTarayici, "Opera 8", 1) > 0 Then
          strKullanilanTarayici = "Opera 8"
 
     ElseIf InStr(1, strTarayici, "Opera", 1) > 0 Then
          strKullanilanTarayici = "Opera"
 
     '// AOL
     ElseIf inStr(1, strTarayici, "AOL 3", 1) > 0 Then
          strKullanilanTarayici = "AOL 3"
 
     ElseIf inStr(1, strTarayici, "AOL 4", 1) > 0 Then
          strKullanilanTarayici = "AOL 4"
 
     ElseIf inStr(1, strTarayici, "AOL 5", 1) > 0 Then
          strKullanilanTarayici = "AOL 5"
 
     ElseIf inStr(1, strTarayici, "AOL 6", 1) > 0 Then
          strKullanilanTarayici = "AOL 6"
 
     ElseIf inStr(1, strTarayici, "AOL 7", 1) > 0 Then
          strKullanilanTarayici = "AOL 7"
 
     ElseIf inStr(1, strTarayici, "AOL 8", 1) > 0 Then
          strKullanilanTarayici = "AOL 8"
 
     ElseIf inStr(1, strTarayici, "AOL 9", 1) > 0 Then
          strKullanilanTarayici = "AOL 9"
 
     ElseIf inStr(1, strTarayici, "AOL", 1) > 0 Then
          strKullanilanTarayici = "AOL"
 
 
     '// Konqueror
     ElseIf inStr(1, strTarayici, "Konqueror", 1) > 0 Then
          strKullanilanTarayici = "Konqueror"
 
 
     '// EudoraWeb
     ElseIf inStr(1, strTarayici, "EudoraWeb", 1) > 0 Then
          strKullanilanTarayici = "EudoraWeb"
 
 
     '// Dreamcast
     ElseIf inStr(1, strTarayici, "Dreamcast", 1) > 0 Then
          strKullanilanTarayici = "Dreamcast"
     
 
     '// Safari
     ElseIf inStr(1, strTarayici, "Safari", 1) > 0 Then
          strKullanilanTarayici = "Safari"
     
 
     '// Lynx
     ElseIf inStr(1, strTarayici, "Lynx", 1) > 0 Then
          strKullanilanTarayici = "Lynx"
     
 
     '// ICE
     ElseIf inStr(1, strTarayici, "ICE", 1) > 0 Then
          strKullanilanTarayici = "ICE"
     
 
     '// iCab 
     ElseIf inStr(1, strTarayici, "iCab", 1) > 0 Then
          strKullanilanTarayici = "iCab"
          
 
     '// HotJava 
     ElseIf inStr(1, strTarayici, "Sun", 1) > 0 AND inStr(1, strTarayici, "Mozilla/3", 1) > 0 Then
          strKullanilanTarayici = "HotJava"
     
 
     '// Galeon 
     ElseIf inStr(1, strTarayici, "Galeon", 1) > 0 Then
          strKullanilanTarayici = "Galeon"
          
 
     '// Epiphany 
     ElseIf inStr(1, strTarayici, "Epiphany", 1) > 0 Then
          strKullanilanTarayici = "Epiphany"
     
 
     '// DocZilla 
     ElseIf inStr(1, strTarayici, "DocZilla", 1) > 0 Then
          strKullanilanTarayici = "DocZilla"
     
 
     '// Camino 
     ElseIf inStr(1, strTarayici, "Chimera", 1) > 0 OR inStr(1, strTarayici, "Camino", 1) > 0 Then
          strKullanilanTarayici = "Camino"
     
 
     '// Dillo 
     ElseIf inStr(1, strTarayici, "Dillo", 1) > 0 Then
          strKullanilanTarayici = "Dillo"
          
 
     '// amaya 
     ElseIf inStr(1, strTarayici, "amaya", 1) > 0 Then
          strKullanilanTarayici = "Amaya"
          
 
     '// NetCaptor 
     ElseIf inStr(1, strTarayici, "NetCaptor", 1) > 0 Then
          strKullanilanTarayici = "NetCaptor"
          
 
     '// Arama Motoru Robotlari     
          
          
     '// LookSmart 
     ElseIf inStr(1, strTarayici, "ZyBorg", 1) > 0 Then
          strKullanilanTarayici = "LookSmart"          
     
 
     '// Googlebot 
     ElseIf inStr(1, strTarayici, "Googlebot", 1) > 0 Then
          strKullanilanTarayici = "Googlebot"
          
 
     '// MSN 
     ElseIf inStr(1, strTarayici, "msnbot", 1) > 0 Then
          strKullanilanTarayici = "MSN"
          
 
     '// inktomi
     ElseIf inStr(1, strTarayici, "slurp", 1) > 0 Then
          strKullanilanTarayici = "Inktomi"
     
 
     '// AltaVista
     ElseIf inStr(1, strTarayici, "Scooter", 1) > 0 Then
          strKullanilanTarayici = "AltaVista"
     
 
     '// DMOZ
     ElseIf inStr(1, strTarayici, "Robozilla", 1) > 0 Then
          strKullanilanTarayici = "DMOZ"
          
 
     '// Ask Jeeves
     ElseIf inStr(1, strTarayici, "Ask Jeeves", 1) > 0 OR inStr(1, strTarayici, "Ask+Jeeves", 1) > 0 Then
          strKullanilanTarayici = "Ask Jeeves"
          
 
     '// Lycos
     ElseIf inStr(1, strTarayici, "lycos", 1) > 0 Then
          strKullanilanTarayici = "Lycos"
          
 
     '// Excite
     ElseIf inStr(1, strTarayici, "ArchitextSpider", 1) > 0 Then
          strKullanilanTarayici = "Excite"
          
 
     '// Northernlight
     ElseIf inStr(1, strTarayici, "Gulliver", 1) > 0 Then
          strKullanilanTarayici = "Northernlight"
     
 
     '// AllTheWeb
     ElseIf inStr(1, strTarayici, "crawler@fast", 1) > 0 Then
          strKullanilanTarayici = "AllTheWeb"
          
 
     '// Turnitin
     ElseIf inStr(1, strTarayici, "TurnitinBot", 1) > 0 Then
          strKullanilanTarayici = "Turnitin"
          
 
     '// InternetSeer
     ElseIf inStr(1, strTarayici, "internetseer", 1) > 0 Then
          strKullanilanTarayici = "InternetSeer"
          
 
     '// NameProtect Inc.
     ElseIf inStr(1, strTarayici, "nameprotect", 1) > 0 Then
          strKullanilanTarayici = "NameProtect"
          
 
     '// PhpDig
     ElseIf inStr(1, strTarayici, "PhpDig", 1) > 0 Then
          strKullanilanTarayici = "PhpDig"
          
 
     '// Rambler
     ElseIf inStr(1, strTarayici, "StackRambler", 1) > 0 Then
          strKullanilanTarayici = "Rambler"
          
 
     '// UbiCrawler
     ElseIf inStr(1, strTarayici, "UbiCrawler", 1) > 0 Then
          strKullanilanTarayici = "UbiCrawler"
          
 
     '// entireweb
     ElseIf inStr(1, strTarayici, "Speedy+Spider", 1) > 0 Then
          strKullanilanTarayici = "entireweb"
          
 
     '// Alexa.com
     ElseIf inStr(1, strTarayici, "ia_archiver", 1) > 0 Then
          strKullanilanTarayici = "Alexa"
     
 
     '// Arianna/Libero
     ElseIf inStr(1, strTarayici, "arianna.libero.it", 1) > 0 Then
          strKullanilanTarayici = "Arianna/Libero"
               
     
          
     '// Internet Explorer
     ElseIf inStr(1, strTarayici, "MSIE 8", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"

     ElseIf inStr(1, strTarayici, "MSIE 7", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
     ElseIf inStr(1, strTarayici, "MSIE 6", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
     ElseIf inStr(1, strTarayici, "MSIE 5", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
     ElseIf inStr(1, strTarayici, "MSIE 4", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
     ElseIf inStr(1, strTarayici, "MSIE 3", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
     ElseIf inStr(1, strTarayici, "MSIE 2", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
     ElseIf inStr(1, strTarayici, "MSIE 1", 1) > 0 Then
          strKullanilanTarayici = "Internet explorer"
 
 
     '// Pocket Internet Explorer
     ElseIf inStr(1, strTarayici, "MSPIE 1", 1) > 0 Then
          strKullanilanTarayici = "Pocket IE 1"
 
     ElseIf inStr(1, strTarayici, "MSPIE 1", 1) > 0 Then
          strKullanilanTarayici = "Pocket IE 2"
 
     
     '// Mozilla Firefox
     ElseIf inStr(1, strTarayici, "Gecko", 1) > 0 AND inStr(1, strTarayici, "Firefox", 1) > 0 Then
          strKullanilanTarayici = "Mozilla Firefox"
     
 
     '// Mozilla Firebird
     ElseIf inStr(1, strTarayici, "Gecko", 1) > 0 AND inStr(1, strTarayici, "Firebird", 1) > 0 Then
          strKullanilanTarayici = "Mozilla Firebird"
     
 
     '// Mozilla
     ElseIf inStr(1, strTarayici, "Gecko", 1) > 0 AND inStr(1, strTarayici, "rv:2", 1) > 0 AND inStr(1, strTarayici, "Netscape", 1) = 0 Then
          strKullanilanTarayici = "Mozilla 2"
 
     ElseIf inStr(1, strTarayici, "Gecko", 1) > 0 AND inStr(1, strTarayici, "rv:1", 1) > 0 AND inStr(1, strTarayici, "Netscape", 1) = 0 Then
          strKullanilanTarayici = "Mozilla 1"
 
     ElseIf inStr(1, strTarayici, "Gecko", 1) > 0 AND inStr(1, strTarayici, "rv:0", 1) > 0 AND inStr(1, strTarayici, "Netscape", 1) = 0 Then
          strKullanilanTarayici = "Mozilla"
 
 
 
     '// Netscape
     ElseIf inStr(1, strTarayici, "Netscape/8", 1) > 0 Then
          strKullanilanTarayici = "Netscape 8"
 
     ElseIf inStr(1, strTarayici, "Netscape/7", 1) > 0 Then
          strKullanilanTarayici = "Netscape 7"
 
     ElseIf inStr(1, strTarayici, "Netscape6", 1) > 0 Then
          strKullanilanTarayici = "Netscape 6"
 
     ElseIf inStr(1, strTarayici, "Mozilla/4", 1) > 0 Then
          strKullanilanTarayici = "Netscape 4"
 
     ElseIf inStr(1, strTarayici, "Mozilla/3", 1) > 0 Then
          strKullanilanTarayici = "Netscape 3"
 
     ElseIf inStr(1, strTarayici, "Mozilla/2", 1) > 0 Then
          strKullanilanTarayici = "Netscape 2"
 
     ElseIf inStr(1, strTarayici, "Mozilla/1", 1) > 0 Then
          strKullanilanTarayici = "Netscape 1"

     ElseIf InStr(strTarayici, "Chrome") Then
          strKullanilanTarayici = "Chrome"
     '// Hiçbiri Degilse
     Else
          strKullanilanTarayici = "Bilinmeyen Taracýyý Türü"
     End If
 
     BrowserType = strKullanilanTarayici

End Function
%> 