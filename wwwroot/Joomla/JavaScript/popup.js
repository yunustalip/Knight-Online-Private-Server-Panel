<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->

// Tab Menü Yapma
function showtab(slot,item) {
	// document.getElementById("t"+ slot +""+ item).blur(); //
	for (var i=1; i<= 10; i++){
	  if (document.getElementById("tab" + slot + i)) {
		document.getElementById("tab"+ slot + i).style.display = "none";
		document.getElementById("t"+ slot +""+ i).className = "";
	 }
	}
	document.getElementById("t"+ slot +""+ item).className = "current";
	document.getElementById("tab"+ slot + item).style.display = "block";
}
// Tab menü sonu

var code='';
var ps728 = 0;
var ps160 = 0;
var ps160id = 0;
var now = new Date();
var externalPS160 = 0;
var nIndex = now.getTime();


/**
* GET-URL Parameter mit Javascript verarbeiten
*
*/
function getQueryVariable(variable) {
   var query = window.location.search.substring(1);
   var vars = query.split("&");
   for (var i=0;i<vars.length;i++) {
      var pair = vars[i].split("=");
      if (pair[0] == variable) return pair[1];
   }
   return false;
}

/**
* This function decodes the any string
* that's been encoded using URL encoding technique
*/
function URLDecode(psEncodeString)
{
  // Create a regular expression to search all +s in the string
  var lsRegExp = /\+/g;
  // Return the decoded string
  return unescape(String(psEncodeString).replace(lsRegExp, " "));
}

/** 
* Referer Obje
*/
function user_referrer()
{
   this.host = "";
   this.query = "";
}

/**
* Arama sitesinden gelen Referer
*/
function ReferrerDecode()
{
   if (document.referrer&&document.referrer!="")
      {
         var result = new user_referrer();
         var ref = document.referrer;

         var vars = ref.split("&");
         for (var i=0;i<vars.length;i++) {
            var pair = vars[i].split("=");
            if (pair[0] == "q") result.query = pair[1];
         }

         result.host = ref.match(/http:\/\/\S*\//);
         return result;
      }
   else return false;
}


/* Hover Menu */

sfHover = function() {
   var sfEls = document.getElementById("he-v1-nav").getElementsByTagName("LI");
   for (var i=0; i<sfEls.length; i++) {
      sfEls[i].onmouseover=function() {
         this.className+=" sfhover";
      }
      sfEls[i].onmouseout=function() {
         this.className=this.className.replace(new RegExp(" sfhover\\b"), "");
      }
   }
}

//if (window.attachEvent) window.attachEvent("onload", sfHover);

/* Arama Kelimeleri yerleþtirir */

var ana_arama_kelimesi = 'Arama';
function arama_kelimesi_yerlestir(fieldId)
{
   var ref;
   var q = getQueryVariable("q");
   var pfad = window.location.pathname;
   var isSuche = pfad.match(/arama.asp/);
   var searches = /(google)/;
   ref = ReferrerDecode()
   if (typeof(ref) == "object"){
      if(searches.test(ref.host)){
         document.getElementById(fieldId).value = URLDecode(ref.query);
      } else 
      {
         if ( (q == "") || !isSuche )
         {
            if (ana_arama_kelimesi == 'Arama')
            {
               try {
	          var arr_size = arama_kelimeleri.length;
                  var random_num = (Math.round((Math.random()*arr_size)-1));
	          ana_arama_kelimesi =  arama_kelimeleri[random_num];
                } catch (e) {
                   document.getElementById(fieldId).value = 'Arama';	
                }
            }
            document.getElementById(fieldId).value = ana_arama_kelimesi;
            } else
         {
            document.getElementById(fieldId).value = URLDecode(q);
         }
      }
   }
   return true;
}

var pmwin = window;

function chgUrun(strId) {
	var sidm = strId;
	document.getElementById("urb"+sidm).style.display = 'none';
	document.getElementById("umb"+sidm).style.display = 'block';
	document.getElementById("urm"+sidm).style.display = 'none';
	document.getElementById("umm"+sidm).style.display = 'block';
}

function chgUrunb(strId) {
	var sidm = strId;
	document.getElementById('urb'+sidm).style.display = 'block';
	document.getElementById('umb'+sidm).style.display = 'none';
	document.getElementById('urm'+sidm).style.display = 'block';
	document.getElementById('umm'+sidm).style.display = 'none';
}

function pm_frame() {
	if (!pmwin.closed)
		pmwin.focus();
	pmwin=window.open("/forum/pop_pm.asp?action=frame", "pm", "width=600,height=440,scrollbars=yes,resizable=yes");
}

function openWindow(url) {
	popupWin = window.open(url,'new_page','width=600,height=400')
}
function openWindow2(url) {
	popupWin = window.open(url,'new_page','width=400,height=450')
}
function openWindow3(url) {
	popupWin = window.open(url,'new_page','width=400,height=450,scrollbars=yes')
}
function openWindow4(url) {
	popupWin = window.open(url,'new_page','width=400,height=525')
}
function openWindow5(url) {
	popupWin = window.open(url,'new_page','width=450,height=525,scrollbars=yes,toolbars=yes,menubar=yes,resizable=yes')
}
function openWindow6(url) {
	popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes')
}
function openWindowHelp(url) {
	popupWin = window.open(url,'new_page','width=470,height=200,scrollbars=yes')
}

function konugit(konuid) {
		var newLoc = "/konu.asp?id="+konuid
		document.location.href = newLoc;
}

function linkGit(strURL) {
		var newLoc = strURL
		document.location.href = newLoc;
}

function hideDiv(divID) {
    document.getElementById(''+divID+'').style.display = 'none';
}

function displayDiv(divID)  {
    document.getElementById(''+divID+'').style.display = 'block';
}
var crTab = 1;
var crOld = 1;

function chTab(ttb) {
		 	document.getElementById("mtab"+crOld).style.display= "none";
		 	document.getElementById("mtab"+ttb).style.display= "";
		 	document.getElementById("tab"+crOld).className = "";
		 	document.getElementById("tab"+ttb).className = "act";
		 	crOld = ttb;
}	 
function oyOver(varOy) {
	var newOy = varOy;
	document.getElementById("yildizlar").className = 'yildizo'+varOy;
}

function oyOut(varOy) {
	var newOy = varOy;
	document.getElementById("yildizlar").className = 'yildiz'+varOy;
}
function oyOut2(varOy) {
	var newOy = varOy;
	document.getElementById("yildizlar").className = 'yildizo'+varOy;
}
function yildizoyla(varStryildiz) {
	
	
	document.FormOylama.action = document.location.href;
	document.FormOylama.yildizoy.value = varStryildiz; 
	document.FormOylama.submit();
	
}

function voteYorum(varId,varNot) {
	document.yorumoy.action = document.location.href;
	document.yorumoy.yorumid.value = varId;
	document.yorumoy.notid.value = varNot;
	document.yorumoy.submit();
}

function yorumuGonder() {
	document.getElementById("yorumgonderbut").style.visibility = 'hidden';
	document.yrmgonder.action = document.location.href;
	document.yrmgonder.submit();
}

function gitAnchor(varURL,varDiyez) {
	var newURL = varURL;
	var newDiyez = varDiyez
	ayirgac = newURL.split("#");
	var yeniURL = ayirgac[0];
	document.location = yeniURL + '#' + varDiyez;
}
function txtvidOver(intNum) {
	document.getElementById('videob').innerHTML = videobaslik[intNum];
	document.getElementById('videom').innerHTML = videoaciklama[intNum];
	if (intNum != 0) {
		document.getElementById('vid1').className = 'viddiv';
	} 
}
function txtgalOver(intNum) {
	document.getElementById('gbaslik').innerHTML = galeribaslik[intNum];
	document.getElementById('gaciklama').innerHTML = galeriaciklama[intNum];
	if (intNum != 0) {
		document.getElementById('anagal1').className = 'galeridiv';
	} 
}
function txtgalOut(intNum) {
	document.getElementById('gbaslik').innerHTML = galeribaslik[0];
	document.getElementById('gaciklama').innerHTML = galeriaciklama[0];
}

function txtgalOvers(intNum) {
	document.getElementById('gbasliks').innerHTML = galeribasliks[intNum];
	document.getElementById('gaciklamas').innerHTML = galeriaciklamas[intNum];
	if (intNum != 0) {
		document.getElementById('anagal1s').className = 'galeridiv';
	} 
}
function txtgalOuts(intNum) {
	document.getElementById('gbasliks').innerHTML = galeribasliks[0];
	document.getElementById('gaciklamas').innerHTML = galeriaciklamas[0];
}


function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).length)
	    return str;
	else
	    return String(str).substring(0,n);
}

function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}
function addFacebook(strThumb, strURL, strTitle, strDesc) {
	window.open('http://www.facebook.com/sharer.php?s=100&p[medium]=100&p[title]='+strTitle+'&p[images][0]='+strThumb+'&p[url]='+strURL+'&p[summary]='+strDesc+'','sharer','toolbar=0,status=0,width=626,height=436');
}

function Bookmarkekle (strURL, strTitle) {
	if (window.sidebar) { // Mozilla Firefox Bookmark
		window.sidebar.addPanel(strTitle, strURL,"");
	} else if( window.external ) { // IE Favorite
		window.external.AddFavorite(strURL, strTitle); }
	else if(window.opera && window.print) { // Opera Hotlist
		var elem = document.createElement('a');
		elem.setAttribute('href',url);
		elem.setAttribute('title',title);
		elem.setAttribute('rel','sidebar');
		elem.click();
 }
}
function addDelicious (strTitle, strURL, strDesc, strKeyword) {
	window.open('http://del.icio.us/post?v=4&noui&jump=close&url='+strURL+'&title='+strTitle+'&notes='+strDesc+'&tags='+strKeyword, 'delicious','toolbar=no,width=700,height=400');
}

function addFacebook (strThumb, strURL, strTitle, strDesc) {
	window.open('http://www.facebook.com/sharer.php?s=100&p[medium]=100&p[title]='+strTitle+'&p[images][0]='+strThumb+'&p[url]='+strURL+'&p[summary]='+strDesc+'','sharer','toolbar=0,status=0,width=626,height=436');
}

function addGBookmark (strTitle, strURL, strDesc, strKeyword) {
	window.open('http://www.google.com/bookmarks/mark?op=edit&output=popup&bkmk='+strURL+'&title='+strTitle+'&labels='+strKeyword+'&annotation='+strDesc, 'googlebookmark','toolbar=no,width=700,height=500');
}
function addDigg (strTitle, strURL, strDesc) {
	window.open('http://digg.com/submit?phase=2&url='+document.location.href+'&title='+strTitle, 'digg','scrollbars=yes,toolbar=no,width=760,height=500');
}
function addYahoo (strTitle, strURL, strDesc, strKeyword) {
	window.open('http://myweb2.search.yahoo.com/myresults/bookmarklet?u='+strURL+'&t='+strTitle2+'&tag='+strKeyword+'&d='+strDesc, 'yahoo','toolbar=no,width=700,height=400');
}
function addTusul (strURL) {
	window.open('http://www.tusul.com/submit.php?url='+strURL, 'tusul','toolbar=no,width=700,height=400');
}
function addTechnorati (strUL) {
	window.open('http://www.technorati.com/faves?add='+strURL, 'technorati','scrollbars=yes,toolbar=no,width=800,height=500');
}
function addSpurl (strTitle, strURL) {
	window.open('http://www.spurl.net/spurl.php?url='+strURL+'&title='+strTitle,'spurl','toolbar=no,width=700,height=400');
}
function resimBuyut (strImage, strXb, strYb) {
	window.open('/araclar/resimgoster.html?resim='+strImage,'Buyuk','toolbar=no,width='+ strXb +',height='+ strYb +',scrollbars=yes');
}

function ccn(strId,strClass) {
	var cId = document.getElementById(strId);
	cId.className = strClass;
}

function galeriBuyuk (strURL, strXb, strYb) {
	window.open(strURL,'galeri','toolbars=no,width='+ (strXb+20) +',height='+ (strYb+70)+',scrollbars=yes');
}
function setVisible(varId) {
	document.getElementById(varId).style.display = 'block';
}

function setInVisible(varId) {
	document.getElementById(varId).style.display = 'none';
}
function hideFloat() {
	 var obj = document.getElementById("float");
	 obj.style.display = 'none';
}



function log_out(B){var A=document.getElementsByTagName("html")[0];A.style.filter="progid:DXImageTransform.Microsoft.BasicImage(grayscale=1)";if(confirm(B)){return true}else{A.style.filter="";return false}}
	
var xScroll, yScroll, timerPoll, timerRedirect, timerClock, rotatorz;


function initRedirect()
{
if (typeof document.body.scrollTop != "undefined"){ //IE,NS7,Moz
    xScroll = document.body.scrollLeft;
    yScroll = document.body.scrollTop;

    clearInterval(timerPoll); //stop polling scroll move
    clearInterval(timerRedirect); //stop timed redirect

    timerPoll = setInterval("pollActivity()",1); //poll scrolling
    timerRedirect = setInterval("document.location.href=document.location.href",300000); //set timed redirect
  }
  else if (typeof window.pageYOffset != "undefined"){ //other browsers that support pageYOffset/pageXOffset instead
    xScroll = window.pageXOffset;
    yScroll = window.pageYOffset;

    clearInterval(timerPoll); //stop polling scroll move
    clearInterval(timerRedirect); //stop timed redirect

    timerPoll = setInterval("pollActivity()",1); //poll scrolling
    timerRedirect = setInterval("document.location.href=document.location.href",300000); //set timed redirect

  
  }
  
}
function pollActivity(){
  if ((typeof document.body.scrollTop != "undefined" && (xScroll!=document.body.scrollLeft || yScroll!=document.body.scrollTop)) //IE/NS7/Moz
   ||
   (typeof window.pageYOffset != "undefined" && (xScroll!=window.pageXOffset || yScroll!=window.pageYOffset))) { //other browsers
      initRedirect(); //reset polling scroll position
  }
}

var i=0;
var oldtitle=0;
var tmp=0;


function rtstart()
{
   // setTimeout('gofwds(+1);',6000);
}

function rotClear() {
	clearInterval(rotatorz);
}
function rotCont() {
	rotatorz = setInterval('gofwds(+1)',6000);
}
function mshow(jj)
{
		if (document.getElementById("rtresim"+jj).src == "") {
			document.getElementById("rtresim"+jj).src = rtresim[jj];
		}
		document.getElementById("konursm"+oldtitle).style.display= "none";
    document.getElementById("konursm"+jj).style.display= "block";
    document.getElementById("secim"+oldtitle+"").className="secimli";
    document.getElementById("secim"+jj+"").className="secimlihover";
    oldtitle=jj;
    clearInterval(rotatorz);
}

function gofwd(kk)
{
    kk=oldtitle+kk;
    if(kk>9) kk=0;
    if(kk<0) kk=9;
    document.getElementById("konursm"+oldtitle).style.display= "none";
    document.getElementById("konursm"+kk).style.display= "";
    document.getElementById("secim"+oldtitle+"").className="secimli";
    document.getElementById("secim"+kk+"").className="secimlihover";
    oldtitle=kk;
    clearInterval(rotatorz);
    
}

function mshow2(jj)
{
	clearInterval(rotatorz);
	rotatorz = setInterval('gofwds(+1)',6000);
}

function gofwds(kk)
{
    clearInterval(rotatorz);
    kk=oldtitle+kk;
    if(kk>9) kk=0;
    if(kk<0) kk=9;
    document.getElementById("konursm"+oldtitle).style.display= "none";
    document.getElementById("konursm"+kk).style.display= "";
    document.getElementById("secim"+oldtitle+"").className="secimli";
    document.getElementById("secim"+kk+"").className="secimlihover";
    oldtitle=kk;
   	rotatorz = setInterval('gofwds(+1)',6000);
    
}

function rtbut(ll)
{
	if (ll==1) {
		document.getElementById("rt_rw").src = "/images/as/rt_rw_in.gif";
	} else {
		document.getElementById("rt_fw").src = "/images/as/rt_fw_in.gif";
	}
	clearInterval(rotatorz);
}


function rtbutb(ll)
{
	if (ll==1) {
		document.getElementById("rt_rw").src = "/images/as/rt_rw_out.gif";
	} else {
		document.getElementById("rt_fw").src = "/images/as/rt_fw_out.gif";
	}
	clearInterval(rotatorz);
	rotatorz = setInterval('gofwds(+1)',6000);
}

function closeMPU(reklamId) {
	document.getElementById(reklamId).style.display = 'none';
	document.getElementById(reklamId).style.visibility = 'hidden';
}	


function trim(stringToTrim) {
	return stringToTrim.replace(/^\s+|\s+$/g,"");
}

function saklagoster(strTabId) {
	var itemId = document.getElementById(""+strTabId);
	var curDisplay = document.getElementById(""+strTabId).style.display;
	
	if (curDisplay == 'none') {
		itemId.style.display = 'block';	
	}
	else {
		itemId.style.display = 'none';
	}
}

function checkie6() {
	var agent = navigator.userAgent.toLowerCase();
	if (agent.indexOf("msie 6") != -1) {
		ie6 = 1;
	} else {
		ie6 = 0;
	}
}
function scrollDiv(varDivId) {
	var el = document.getElementById(varDivId);
	el.style.top = document.body.scrollTop + 'px';	
}

function el_move(){
if(ie4){ydiff=el_height_start-document.body.scrollTop; xdiff=el_width_start-document.body.scrollLeft}
else{ydiff=el_height_start-pageYOffset; xdiff=el_width_start-pageXOffset}
if(ydiff!=0){movey=Math.round(ydiff/10);el_height_start-=movey}
if(xdiff!=0){movex=Math.round(xdiff/10);el_width_start-=movex}
if(ns4){document.layers.element.top=el_height_start+el_height;document.layers.element.right=el_width_start+el_width}
if(ie4){document.all.element.style.top=el_height_start+el_height;document.all.element.style.right=el_width_start+el_width}
if(ns6){document.getElementById("element").style.top=el_height_start+el_height;document.getElementById("element").style.right=el_width_start+el_width}
}

function movereklam(evt) {
				var divId = "bb160";
				var newDiv = document.getElementById(divId).style;
				newDiv.position = 'absolute';
				var newTop = parseInt(document.documentElement.scrollTop);
				var maxLength = parseInt(document.body.scrollHeight);
				
				var newToplam = parseInt(maxLength - parseInt(document.getElementById(divId).offsetHeight));
				if (newTop < newToplam) {
					newDiv.top = newTop + 4 + 'px';
			 }
				
}


function getElementsByClassName(strClass, strTag, objContElm) {
  strTag = strTag || "*";
  objContElm = objContElm || document;
  var objColl = objContElm.getElementsByTagName(strTag);
  if (!objColl.length &&  strTag == "*" &&  objContElm.all) objColl = objContElm.all;
  var arr = new Array();
  var delim = strClass.indexOf('|') != -1  ? '|' : ' ';
  var arrClass = strClass.split(delim);
  for (var i = 0, j = objColl.length; i < j; i++) {
    var arrObjClass = objColl[i].className.split(' ');
    if (delim == ' ' && arrClass.length > arrObjClass.length) continue;
    var c = 0;
    comparisonLoop:
    for (var k = 0, l = arrObjClass.length; k < l; k++) {
      for (var m = 0, n = arrClass.length; m < n; m++) {
        if (arrClass[m] == arrObjClass[k]) c++;
        if (( delim == '|' && c == 1) || (delim == ' ' && c == arrClass.length)) {
          arr.push(objColl[i]);
          break comparisonLoop;
        }
      }
    }
  }
  return arr;
}


function clickMenu(strObj) {
	
	if (strObj.className != "sfhover") 
	{ 
		strObj.className = "sfhover"; 
	} 
	else 
	{ 
		strObj.className= "";
	};	
}