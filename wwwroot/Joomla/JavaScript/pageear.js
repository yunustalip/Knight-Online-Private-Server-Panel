/********************************************************************************************
* PageEar advertising CornerAd by Webpicasso Media
* Leave copyright notice.  
*
* @copyright www.webpicasso.de
* @author    christian harz <pagepeel-at-webpicasso.de>
*********************************************************************************************/
  

/*
 *  Konfiguration
 */ 
 
var pagearSmallImg = 'modules/pageear_s.jpg'; // URL zum kleinen Bild
var pagearSmallSwf = 'modules/pageear_s.swf'; // URL zum thumb.swf

var pagearBigImg = 'modules/pageear_b.jpg'; // URL zum groﬂen Bild
var pagearBigSwf = 'modules/pageear_b.swf'; // URL zum big.swf
 
var mirror = 'true'; // Bild spiegelt sich in der aufgeschlagenen Ecke ( true | false )
var pageearColor = 'ffffff';  // Farbe der aufgeschlagenen Ecke wenn mirror false ist
var jumpTo = 'http://www.joomlasp.com/' // Bura reklamin hedefi. URLsi
var openLink = 'self'; // ÷ffnet den link im neuen Fenster (new) oder im selben (self)

/*
 *  Ab hier nichts mehr ‰ndern 
 */ 

// Flash check vars
var requiredMajorVersion = 6;
var requiredMinorVersion = 0;
var requiredRevision = 0;

// Copyright
var copyright = 'Webpicasso Media, www.webpicasso.de';
 
var thumbWidth  = 100;
var thumbHeight = 100;

var bigWidth  = 500;
var bigHeight = 500;

var queryParams = 'pagearSmallImg='+escape(pagearSmallImg); 
queryParams += '&pagearBigImg='+escape(pagearBigImg); 
queryParams += '&pageearColor='+pageearColor; 
queryParams += '&jumpTo='+escape(jumpTo); 
queryParams += '&openLink='+escape(openLink); 
queryParams += '&mirror='+escape(mirror); 
queryParams += '&copyright='+escape(copyright); 
 
function openPeel(){
	document.getElementById('bigDiv').style.top = '0px';
	document.getElementById('thumbDiv').style.top = '-1000px';
}

function closePeel(){
	document.getElementById("thumbDiv").style.top = "0px";
	document.getElementById("bigDiv").style.top = "-1000px";
}

function writeObjects () { 
    
    // Get installed flashversion
    var hasReqestedVersion = DetectFlashVer(requiredMajorVersion, requiredMinorVersion, requiredRevision);

    // Write div layer for big swf
    document.write('<div id="bigDiv" style="position:absolute;width:'+ bigWidth +'px;height:'+ bigHeight +'px;z-index:9999;right:-1px;top:-1000px;">');    	
    
    // Check if flash exists/ version matched
    if (hasReqestedVersion) {    	
    	AC_FL_RunContent(
    				"src", pagearBigSwf+'?'+ queryParams,
    				"width", bigWidth,
    				"height", bigHeight,
    				"align", "middle",
    				"id", "bigSwf",
    				"quality", "high",
    				"bgcolor", "#FFFFFF",
    				"name", "bigSwf",
    				"wmode", "transparent",
    				"allowScriptAccess","sameDomain",
    				"type", "application/x-shockwave-flash",
    				'codebase', 'http://fpdownload.macromedia.com/get/flashplayer/current/swflash.cab',
    				"pluginspage", "http://www.adobe.com/go/getflashplayer"
    	);
    } else {  // otherwise do nothing or write message ...    	 
    	document.write(alternateContent);  // non-flash content
    } 
    // Close div layer for big swf
    document.write('</div>'); 
    
    // Write div layer for small swf
    document.write('<div id="thumbDiv" style="position:absolute;width:'+ thumbWidth +'px;height:'+ thumbHeight +'px;z-index:9999;right:-1px;top:0px;">');
    
    // Check if flash exists/ version matched
    if (hasReqestedVersion) {    	
    	AC_FL_RunContent(
    				"src", pagearSmallSwf+'?'+ queryParams,
    				"width", thumbWidth,
    				"height", thumbHeight,
    				"align", "middle",
    				"id", "bigSwf",
    				"quality", "high",
    				"bgcolor", "#FFFFFF",
    				"name", "bigSwf",
    				"wmode", "transparent",
    				"allowScriptAccess","sameDomain",
    				"type", "application/x-shockwave-flash",
    				'codebase', 'http://fpdownload.macromedia.com/get/flashplayer/current/swflash.cab',
    				"pluginspage", "http://www.adobe.com/go/getflashplayer"
    	);
    } else {  // otherwise do nothing or write message ...    	 
    	document.write(alternateContent);  // non-flash content
    } 
    document.write('</div>');     
}
