/*======================================================================*\
|| #################################################################### ||
|| #                                                                  # ||
|| #                  FGC WebChat Javascript Functions                # ||
|| # ---------------------------------------------------------------- # ||
|| #    Copyright ©2005 FGC Website Designers. All Rights Reserved.   # ||
|| # This file may not be redistributed in whole or significant part. # ||
|| #                                                                  # ||
|| #     http://www.fgcportal.com | http://webchat.fgcportal.com      # ||
|| #                                                                  # ||
|| #                       info@fgcportal.com                         # ||
|| #                                                                  # ||
|| #################################################################### ||
\*======================================================================*/

function Trim(trimInput) {

	if(trimInput.length < 1) {
		return "";
	}

	trimInput = RTrim(trimInput);
	trimInput = LTrim(trimInput);

	if (trimInput == "") {
		return "";
	} else {
		return trimInput;
	}
}

function RTrim(trimInput) {

	var iLength = trimInput.length;

	if (iLength < 0) {
		return "";
	}
	
	var wSpace = String.fromCharCode(32);
	var sOutPut = "";
	var iTemp = iLength -1;

	while (iTemp > -1) {
		if (trimInput.charAt(iTemp) != wSpace) {
			sOutPut = trimInput.substring(0, iTemp +1);
			break;
		}
		iTemp = iTemp-1;
	}
	return sOutPut;
}

function LTrim(trimInput) {
	
	var iLength = trimInput.length;
	
	if (iLength < 1) {
		return "";
	}
	
	var wSpace = String.fromCharCode(32);
	var sOutPut = "";
	var iTemp = 0;

	while (iTemp < iLength) {
		if (trimInput.charAt(iTemp) != wSpace) {
			sOutPut = trimInput.substring(iTemp, iLength);
			break;
		}
		iTemp = iTemp + 1;
	}
	return sOutPut;
}

function getObject(objID) {
	
	if (document.getElementById) {
		return document.getElementById(objID);
	} else if (document.all) {
		return document.all[objID];
	} else if (document.layers) {
		return document.layers[objID];
	} else {
		return null;
	}
}

function getFlashMovieObject(objID) {
	
	if (window.document[objID]) {
		return window.document[objID];
	}
	if (navigator.appName.indexOf("Microsoft Internet") == -1) {
		if (document.embeds && document.embeds[objID]) {
			return document.embeds[objID];
		}
	} else {
		return document.getElementById(objID);
	}
}

function isWhitespace(sInput) {
	var whiteSpace = /^\s+$/;
	return (whiteSpace.test(sInput));
}

function isEmpty(sInput) {
	try {		
		if (sInput == null || Trim(sInput).length == 0 || isWhitespace(sInput)) {
			return true;
		} else {
			return false;
		}
	} catch(e) {
		return sInput;
	}
}
