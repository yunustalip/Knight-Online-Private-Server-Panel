<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<!--#include file="md5.asp"-->
<%Session.CodePage=1254
Dim sitesettings,sitesql,sitebaslik,SunucuAdi,IP,drop,expt,war,wartime,menuayar,banner,bannercolor,saat,tarih,nesne,serverip,gosterim,objXmlHttp,verimiz,veriler,onlineu,online,ay,gun,greenc,y,r,menu,ver,warhero,activeusers,toplamaktif,gmlist,onlinegm,cat,evt,anket,sy,toplamoy,topoy,ixx,ason,news,x,humansayi,karussayi,toplamchar,ksy,hsy,humansay,karussay,uch,Effect
Set sitesettings=Conne.Execute("select * from siteayar")
sitebaslik = sitesettings("sitebaslik")
SunucuAdi = sitesettings("SunucuAdi")
IP = sitesettings("IP")
Drop =  sitesettings("Droprate")
Expt = sitesettings("Expt")
war = sitesettings("war")
wartime = sitesettings("wartime")
radyo = sitesettings("radyo")
radyokonum = sitesettings("radyokonum")
banner = sitesettings("banner")
bannercolor = sitesettings("bannercolor")
Effect=Split(sitesettings("Effect"),",")




Saat=time()
Tarih=date()
Dim UserIP
UserIP=Request.ServerVariables("REMOTE_HOST")
Set Nesne = CreateObject("Scripting.FileSystemObject")
If Nesne.FileExists(Server.MapPath("logs/"&tarih&"-"&md5(tarih)&".txt"))=False Then
Dim Yaz
Set Yaz = Nesne.CreateTextFile(Server.MapPath("logs/"&tarih&"-"&md5(tarih)&".txt"),True)
Yaz.Write("IP: "&UserIP&" Tarih: "&Now()&" Giriþ Yaptý.")
Yaz.Close
Set yaz=Nothing
else
Dim Ag
Set Ag = Nesne.OpenTextFile(Server.MapPath("logs/"&tarih&"-"&md5(tarih)&".txt"),8,-2)
Ag.Write(vbcrlf&"IP: "&UserIP&" Tarih: "&Now()&" Giriþ Yaptý.")
Ag.Close
Set Ag=Nothing
End If


Set Onlineu =Conne.Execute("Select Count(strcharid) as toplam From CurrentUser")
Online = Onlineu("toplam")
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" >
<title><%=sitebaslik%></title>
<link rel="shortcut icon" href="imgs/favicon.ico" >
<link rel="stylesheet" type="text/css" href="css/homepage.css">
<link rel="stylesheet" type="text/css" href="css/webstyle.css" >
<link rel="stylesheet" type="text/css" href="css/tssanalyazi.css" >
<link rel="stylesheet" href="css/sIFR-screen.css" type="text/css" media="screen" />
<link rel="stylesheet" href="css/sIFR-print.css" type="text/css" media="print" />
<script defer="defer" type="text/javascript" src="js/pngfix.js"></script>
<script type="text/javascript" src="js/jquery.js"></script>
<script type="text/javascript" src="js/overlib.js"></script>
<script type="text/JavaScript" src="js/hompg.js"></script>
<script type="text/javascript" src="js/jquery.validate.js"></script>
<script type="text/javascript" src="js/sifr.js"></script>
<script type="text/javascript" src="js/sifr-addons.js"></script>
<script src="js/tsvkboard.js" type="text/javascript"></script>
<script src="js/tssanal.js" type="text/javascript"></script>
<script src="js/tssanalyazi.js" type="text/javascript"></script>
<style>
.ylink{
color:#00FF00;
}
div#gmmen{
position:fixed;
bottom:0px;
left:0px;
width:100%;
height:23px;
font-size:10px;
z-index:5;
font-weight:bold;
background-color:#000000;
color:#00FF00;
}
div#notice{
position:fixed;
top:20px;
}
<%If Radyo="1" Then
If RadyoKonum="ust" Then
With Response
.Write"div#radiobar{"&vbcrlf
.Write"position:fixed;"&vbcrlf
.Write"top:0px;"&vbcrlf
.Write"left:0px;"&vbcrlf
.Write"width:100%;"&vbcrlf
.Write"height:60px;"&vbcrlf
.Write"font-size:10px;"&vbcrlf
.Write"z-index:5;"&vbcrlf
.Write"font-weight:bold;"&vbcrlf
.Write"color:#00FF00;"&vbcrlf
.Write"}"&vbcrlf

.Write"* html div#radiobar{"&vbcrlf
.Write"position: absolute !important;"&vbcrlf
.Write"top: expression(((document.documentElement.scrollTop || document.body.scrollTop) + this.offsetHeight-90) + ""px"");"&vbcrlf
.Write"left:0px;"&vbcrlf
.Write"width:100%;"&vbcrlf
.Write"height:60px;"&vbcrlf
.Write"font-size:10px;"&vbcrlf
.Write"z-index:5;"&vbcrlf
.Write"font-weight:bold;"&vbcrlf
.Write"color:#00FF00;"&vbcrlf
.Write"}"&vbcrlf
.Write"* html #player2{"&vbcrlf
.Write"position:relative;top:6px"&vbcrlf
.Write"}"&vbcrlf

End With
ElseIf RadyoKonum="alt" Then
With Response
.Write"div#radiobar{"&vbcrlf
.Write"position:fixed;"&vbcrlf
.Write"bottom:0px;"&vbcrlf
.Write"left:0px;"&vbcrlf
.Write"width:100%;"&vbcrlf
.Write"height:60px;"&vbcrlf
.Write"font-size:10px;"&vbcrlf
.Write"z-index:5;"&vbcrlf
.Write"font-weight:bold;"&vbcrlf
.Write"color:#00FF00;"&vbcrlf
.Write"}"&vbcrlf

.Write"* html div#radiobar{"&vbcrlf
.Write"position: absolute !important;"&vbcrlf
.Write"top: expression(((document.documentElement.scrollTop || document.body.scrollTop) + (document.documentElement.clientHeight || document.body.clientHeight) - this.offsetHeight) + ""px"");"&vbcrlf
.Write"left:0px;"&vbcrlf
.Write"width:100%;"&vbcrlf
.Write"height:60px;"&vbcrlf
.Write"font-size:10px;"&vbcrlf
.Write"z-index:5;"&vbcrlf
.Write"font-weight:bold;"&vbcrlf
.Write"color:#00FF00;"&vbcrlf
.Write"}"&vbcrlf
.Write"* html #player2{"&vbcrlf
.Write"position:relative;top:30px"&vbcrlf
.Write"}"&vbcrlf
End With
End If
End If%>
h1#banner {
	font-size: 50px;
	text-align:center;
}
h1#girisyazi {
	font-size: 58px;
}

* html div#notice{
position: absolute !important;
top: expression(((document.documentElement.scrollTop || document.body.scrollTop) + this.offsetHeight) + "px");
width: expression(document.documentElement.clientWidth || document.body.clientWidth-50+"px");
}
* html div#gmmen{
position: absolute !important;
top: expression(((document.documentElement.scrollTop || document.body.scrollTop) + (document.documentElement.clientHeight || document.body.clientHeight) - this.offsetHeight) + "px");
left:0px;
width: expression(document.documentElement.clientWidth || document.body.clientWidth-100+"px");
height:20px;
font-size:10px;
z-index:5;
font-weight:bold;
background-color:#000;
color:#00FF00;
}

* html a{
cursor:pointer;
}
* html a:hover {
cursor:pointer;
}
* html a:active{
cursor:pointer;
}

* html .ylink{
color:#00FF00;
}
</style>

<script language="JavaScript" type="text/JavaScript">
//<![CDATA[
/* Replacement calls. Please see documentation for more information. */
if(typeof sIFR == "function"){
// This is the preferred "named argument" syntax

sIFR.replaceElement(named({sSelector:"h1#banner", sFlashSrc:"font/UniqueFont.swf", sWmode: "transparent" , sColor:"<%=bannercolor%>",  nPaddingTop:55, nPaddingBottom:0, sFlashVars:"textalign=center&offsetTop=0"}));

sIFR.replaceElement(named({sSelector:"h1#girisyazi", sFlashSrc:"font/ring.swf", sWmode: "transparent" , sColor:"<%=bannercolor%>",  nPaddingTop:0, nPaddingBottom:0, sFlashVars:"textalign=center&offsetTop=0"}));
// This is the older, ordered syntax
};
//]]>


if(location.hash!=""){
openpage(location.hash.substring(1,location.hash.length)+".html")
}

function noticereload(){
$.ajax({
   url: 'notice.asp',
   success: function(ajaxCevap) {
      $('div#notice').fadeOut("slow").fadeIn("slow").html(ajaxCevap);
   }
});
}

setTimeout(noticereload,1000)
setInterval(noticereload,50000)

function loadingx() {
window.status='Sayfa Yükleniyor...'
}
function successx(){
window.status='Yüklendi!'
}
function chngtitle(id){
document.title= id+' > <%=sitebaslik%>'
}
var counter = 30;

function AddZero(rakam)
	{
	return (rakam < 10) ? '0' + rakam : rakam;
	}

function AddZeroMnth(rakam)
	{
	rakam = rakam + 1
	return (rakam < 10) ? '0' + rakam : rakam;
	}

	function timeDiff()	
		{
		var timeDifferense;
		var serverClock = new Date(<%=right(tarih,4)&","&mid(tarih,4,2)&","&mid(tarih,1,2)&","&replace(saat,":",",")%>);
		var clientClock = new Date();
		var serverSeconds;
		var clientSeconds;
	
		timeDiff = clientClock.getTime() - serverClock.getTime() - 500;
		runClock(timeDiff);
		}
	function runClock(timeDiff)
		{

		var now = new Date();
		var newTime;
		newTime = now.getTime() - timeDiff;
		now.setTime(newTime);
		{
			if (counter > 0){
			document.getElementById('clock_area').title = 'Server saati.';
			document.getElementById('clock_area').style.color = '#00FF00';
			}
			counter--;
			document.getElementById("clock_area").innerHTML = 'Server Saati: '+AddZero(now.getHours()) + ':' + AddZero(now.getMinutes()) + ':' + AddZero(now.getSeconds()) ;
		}
		setTimeout('runClock(timeDiff)',1000);
	}

function serverdetail(){
$.ajax({
   url: 'ServerDurum.Asp',
   success: function(ajaxCevap) {
      $('#serverdetay').html(ajaxCevap);
   }
});
}
function serverdetail2(){
$.ajax({
   url: 'ServerDurum.Asp?islem=2',
   success: function(ajaxCevap) {
      $('#serverdetails').html(ajaxCevap);
   }
});
}

setInterval(serverdetail,10000);
setInterval(serverdetail2,20000);

function pageload(name,opt){
if(opt=="1"){
  adres=name
}
else if(opt==null){
   adres=name+".html"
}

$.ajax({
url: adres ,
start:$('div#ortabolum').animate({width:'150px'}, 0).html('<center><br><br><br><br><br><img src=imgs/18-1.gif><br><br><div style="color:#89640B;font-weight:bold">Sayfa Yükleniyor...</div></center>'),
success:function(ajaxCevap) {
$('div#ortabolum').stop().animate({width:'100%'}, 500).html(ajaxCevap);
successx();
}
});

scrollTo(0,50);
}

$(document).ready(function() {
 $(window).bind('load', function()
 {
 
   // resim onyükleme fonksiyonu
   jQuery.preloadImages = function()
   {
   for(var i = 0; i<arguments.length; i++)
   {
     jQuery("<img>").attr("src", arguments[i]);
   }
   };
 
   // yükleme yap
   $.preloadImages("imgs/001.gif", "imgs/002.gif", "imgs/003.gif", "imgs/004.gif", "imgs/005.gif", "imgs/006.gif", "imgs/1.gif", "imgs/2.gif", "imgs/3.gif", "imgs/4.gif", "imgs/5.gif", "imgs/6.gif", "imgs/menubg.gif");
 
 });
});

</script>
<script type="text/javascript" src="Flash/swfobject.js"></script>
<script type="text/javascript">
function cal(url){
var so = new SWFObject('Flash/player.swf','mpl','0','0','9');
so.addVariable('file', 'http://localhost/sounds/'+url);
so.addVariable('title', 'My Video');
so.addVariable('flashvars','&autostart=true&');
so.write('player');
}



</script>
   <script>

	var arr = document.getElementsByTagName('img');
	var l = {};
	for(x in arr)
	{
		var obj = eval('arr["'+x+'"]');
		var loading = document.createElement("img");
			loading.setAttribute("src","bekleme.gif");


		if(isFinite(x)){
			if(obj.getAttribute('loading')=='1'){
				obj.style.visibility = "hidden";
				eval('l.g'+x+' = obj');
				eval('l.f'+x+' = document.createElement("div");');
				eval('l.f'+x+'.appendChild(loading)');
				eval('l.f'+x+'.appendChild(document.createTextNode(" Sorguluyorum !"))');
				eval('l.f'+x+'.setAttribute("id","loading");');
				eval('l.f'+x+'.style.left = "490px"');
				eval('l.f'+x+'.style.top = "325px"');
				eval('document.body.appendChild(l.f'+x+');');
		
				eval('obj.addEventListener("load",function(){ l.g'+x+'.style.visibility="visible"; l.f'+x+'.style.display="none" },false);');
				
			}
		}

	}


       </script>
</head>
<body id="bdy">
<style>
#loading { 
   position:absolute;
  }
       </style>
<%If Radyo="1" Then%>
<div id="radiobar" align="left"><object align="left"
id="player2"
name="player2"
classid="CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6"  
codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701"
standby="Loading Microsoft Windows Media Player components..." 
title="Müzik Player"
type="x-ms-wmp"
width="90%" 
height="60">
      <param name="url" value="">
      <param name="animationatStart" value="true">
      <param name="transparentatStart" value="true">
      <param name="autoStart" value="true">
      <param name="showControls" value="true">
      <param name="volume" value="100">
      <param name="loop" value="true">
      <param name="ShowStatusBar" value="true">
<embed src="" id="player1" name="player1"
border="0" 
type="application/x-mplayer2" 
pluginspage="http://www.microsoft.com/isapi/redir.dll?prd=windows&sbp=mediaplayer&ar=Media&sba=Plugin&" 
width=90%
height=60
volume=0
autostart=1
showstatusbar=1
mute=1></embed>
</object><a onClick="javascript:window.open('MediaList.Asp','MediaList','fullscreen=0,top='+screen.availHeight/4+',left='+screen.availWidth/4+',resizable=1,status=1,scrollbars=1,menubar=0,toolbar=0,height=300,width=500');return false" style="color:#00FF00;">Çalma Listesi<br><img src="RadyoFiles/ListenMusic.gif"></a></div>
<%End If

If war="on" and len(wartime)>0 Then
If DateDiff("n",time,wartime)>0 Then%>
<style type="text/css">

#informationbar{
position: fixed;
left: 0;
width: 100%;
text-indent: 5px;
padding: 5px 0;
background-color: Red;
border-bottom: 1px solid black;
font: bold 12px Verdana;
}

* html #informationbar{ /*IE6 hack*/
position: absolute;
width: expression(document.compatMode=="CSS1Compat"? document.documentElement.clientWidth+"px" : bOdy.clientWidth+"px");
}

</style>
<script type="text/javascript">
function informationbar(){
this.displayfreq="always"
this.content='<a href= "javascript:informationbar.close()"><img src= "imgs/remove.gif" style="width: 14px; height: 14px; float: right; border: 0; margin-right: 5px" /></a>'
}

informationbar.prototype.setContent=function(data) {
this.content=this.content+data
document.write('<div id="informationbar" style="top: -500px">'+this.content+'</div>')
}

informationbar.prototype.animatetoview=function(){
var barinstance=this
if (parseInt(this.barref.style.top)<0){
this.barref.style.top=parseInt(this.barref.style.top)+5+"px"
setTimeout(function(){barinstance.animatetoview()} , 50)
}
else{
if (document.all && !window.XML (''))
this.barref.style.setExpression("top", 'document.compatMode=="CSS1Compat"? document.documentElement.scrollTop+"px" : bOdy.scrollTop+"px"')
else
this.barref.style.top=0
}
}

informationbar.close=function(){
document.getElementById("informationbar").style.display="none"
if (this.displayfreq=="Session")
document.cookie="infobarshown=1;path=/"
}

informationbar.prototype.setfrequency=function(type){
this.displayfreq=type
}

informationbar.prototype.initialize=function(){
if (this.displayfreq=="Session" && document.cookie.indexOf("infobarshown")==-1 || this.displayfreq=="always"){
this.barref=document.getElementById("informationbar")
this.barheight=parseInt(this.barref.offsetHeight)
this.barref.style.top=this.barheight*(-1)+"px"
this.animatetoview()
}
}

window.onunload=function(){
this.barref=null
}
</script>
<script type="text/javascript">
var infobar=new informationbar()
infobar.setContent('<font color=white>Oyuncularýmýzýn Dikkatine Savaþ Baþlamýþtýr! Bitiþ zamaný: <%=wartime%></font>')
infobar.initialize()
</script></div>
<div id="sidebar_container">
<h2 id="sidebar_heading"><span></span></h2>
</div>
</div>
</div>
<%End If
End If
ay=month(date)
gun=day(date)
if ay=4 and gun=23 or ay=5 and gun=19 or ay=8 and gun=30 or ay=10 and gun=29 Then
Response.Write "<div style=""position:absolute;left:-20px;top:0px""><img src=""imgs/cwt.jpg"" width=""200""></div>"
End If
Response.Write "<div id=""player"" style=""height:0px;width:0px"">.</div>"
Dim EffectTime
EffectTime = Request.Cookies("HomePage")("EffectTime")
If EffectTime="" Then EffectTime=Now()-1
If DateDiff("h",CDate(EffectTime),Now())>1 Then
Response.Cookies("HomePage")("EffectTime")=CDbl(Now())
Response.Cookies("HomePage").Expires=Now()+1%>
	<script src="js/jquery.easing.1.3.js" type="text/javascript"></script>  
	<script type="text/javascript">
	cal('position-death.mp3')
	$(document).ready(function acil(){		
	$(".leftcurtain").stop().animate({width:'0px'},{queue:false, duration:10000, easing:'easeInOutExpo'}).remove;
	$(".rightcurtain").stop().animate({width:'0px'},{queue:true, duration:10000, easing:'easeInOutExpo'}).remove;
	$("#giristext").stop().animate({top:document.body.clientHeight/3+'px'},{queue:true, duration:5000, easing:'easeOutBounce'}).hide("slow");
	});


	</script>
	<style type="text/css">
	.leftcurtain{
		width: 50%;
		height: 100%;
		top: 0px;
		left: 0px;
		position: absolute;
		z-index: 7;
	}
	 .rightcurtain{
		width: 51%;
		height: 100%;
		right: 0px;
		top: 0px;
		position: absolute;
		z-index: 8;
	}
	.rightcurtain img, .leftcurtain img{
		width: 100%;
		height: 100%;
	}
</style>
<div class="leftcurtain"><img src="imgs/frontcurtain.jpg"/></div>
<div class="rightcurtain"><img src="imgs/frontcurtain.jpg"/></div>
<div id="giristext" style="position:absolute;left:25%;top:0px;z-index:10;"><h1 id="girisyazi"><%=banner%></h1></div>
<%End If%>


<style>
	.eff{
		left: 240px;
		position:absolute;
		z-index: 6;
	}

	</style>
<div id="notice" style="color:yellow;font-weight:bold;font-size:16;z-index:5;"></div>
<table width="975" border="0" align="center" cellpadding="0" cellspacing="0" id="anatable">
	<tr>
	<td align="left" valign="bottom"><img src="imgs/warriorsol.gif" height="140" width="200"></td>
	<td align="center" valign="middle"><h1 id="banner"><%=banner%></h1></td>
	<td align="right" valign="bottom"><br><br><img src="imgs/warriorsag.gif" height="130" width="178"></td>
	</tr>
	<tr>
	<td ><br>
	<strong><div id="clock_area" align="left" style="margin-left:10;color:#00FF00;" title="Server saati"></div></strong>
	<script>timeDiff();</script><br></td>
	<td align="right" colspan="2"><span id="serverdetay" style="color:#00ff00;font-weight:bold"><%Response.Write "Login Server: "&LoginServerDurum&"&nbsp;&nbsp;Game Server: "&GameServerDurum&"&nbsp;&nbsp;FTP Server: "&FtpServerDurum&"&nbsp;&nbsp;Online Oyuncu: " & Online%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<code style="color:#00ff00;font-weight:bold;font-size:11px">Güvenlik Sistemi Devrede! Flood sayýnýz (<%
if Session("FloodSayi")="" Then
Response.Write "0"
Else
Response.Write Session("FloodSayi")
End If%>/6)</code> <span style="text-decoration:blink;color:#00ff00;">_</span></span></td></tr>
  <tr>
    <td colspan="3" height="25" style="background:url(imgs/content-head.jpg)">
	<span class="style2"><marquee direction="left" width="970"><b>###&nbsp;Notice : &nbsp;<%=sitesettings("duyuru")%>&nbsp;###</b></marquee></span></td>
  </tr>
  <tr>
    <td height="242" valign="top"  colspan="3">
    <table width="975" border="0" cellpadding="0"  cellspacing="0">
  <tr>
        <td width="200" valign="top" >
     <table width="200" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="36" width="200" align="center"><img src="imgs/main_login_title.gif" style="padding-left:1px"></td>
        </tr>
	<tr><td ><div id="menuler" style="background-image:url(imgs/main_login_bg.gif)">
      <table cellpadding="0" cellspacing="0" width="100%" >
	<% set menu=Conne.Execute("select * from menu order by id asc")
	if not menu.eof Then
	do while not menu.eof
	if menu("durum")="1" Then %><tr><td height="20" style="padding-left:12px"><a href="<%=Menu("url")%>" <% if len(menu("click"))>0 Then Response.Write "onclick=""chngtitle('"&menu("menuname")&"');"&menu("click")&""""%> style="display:block" class="link1">&nbsp;<img src="imgs/okz.gif" width="6" height="9" border="0">&nbsp;<%=menu("menuname")%></a></td></tr>
        <%else
	End If
	menu.movenext
	loop
	End If%>
</table>
	</div>
    </td></tr>
      <tr>
        <td height="19" width="185" align="center"  style="background-image:url(imgs/main_login_bottom.gif);background-repeat:no-repeat;padding-left:25px" class="style1"></td>
        </tr>
   </table><br>
<script language="javascript">
function logingiris(){
$.ajax({
   type: 'post',
   url: 'loginok.asp',
   data: $('#loginp').serialize() ,
   start: $('#kullogin').html('<table width="200" cellspacing="0" cellpadding="0" border="0" ><tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px"><img src="imgs/login.gif"></td></tr><tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px"><center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Giriþ Yapýlýyor Lütfen Bekleyin.</center></td></tr><tr> <td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>'),
   success: function(ajaxCevap) {
      $('#kullogin').fadeOut(0).fadeIn("slow").html(ajaxCevap);
   }
});
}


function loging(){
$.ajax({
   url: 'login.asp',
   start:  $('#kullogin').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Lütfen Bekleyin.</center>'),
   success: function(ajaxCevap) {
      $('#kullogin').html(ajaxCevap);
   }
});
}
function logout(){
$.ajax({
   url: 'logout.asp',
   start:  $('#kullogin').html('<table width="200" cellspacing="0" cellpadding="0" border="0" ><tr><td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">Kullanýcý Giriþi</td></tr><tr><td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px"><center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif><br>Çýkýþ Yapýlýyor Lütfen Bekleyin.</center></td></tr><tr> <td height="16" background="imgs/sub_menu_bottom.gif"></td></tr></table>'),
   success: function(ajaxCevap) {
      $('#kullogin').fadeOut(0).fadeIn("slow").html(ajaxCevap);
   }
});
	
}
function logreload(){
$.ajax({
   url: 'login.asp',
   success: function(ajaxCevap) {
      $('#kullogin').html(ajaxCevap);
   }
});
}

window.onload=function(){
setInterval('logreload()',120000);
}

</script>
<div id="sshowimage" style="visibility:hidden;position:absolute;left:150px;top:150px;z-index:10">
	<table border="0" class="turksecurity" cellspacing="0" cellpadding="2">
		<tr>
			<td width="100%">
				<table border="0" width="100%" cellspacing="0" cellpadding="0" height="18px">
					<tr>
						<td id="sdragbar" style="cursor:hand; cursor:pointer" width="100%" onMousedown="sinitializedrag(event)">

							<ilayer width="100%" onSelectStart="return false">
								<layer width="100%" onMouseover="dragswitch=1;if (ns4) sdrag_dropns(sshowimage)" onMouseout="dragswitch=0">
									<font face="Verdana" color="#FFFFFF">
										<strong>
											Sanal Klavye										</strong>									</font>								</layer>
							</ilayer>						</td>
						<td style="cursor:hand">
							<a href="#TurkSecurity.Org" onClick="shidebox();return false">
								<img src="images/kapat.gif" border="0" title="Kapat">							</a>						</td>
					</tr>
				</table>
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td width="100%" bgcolor="#FFFFFF" style="padding:1px">
							<DIV id="keyboard"></DIV>						</td>
					</tr>
				</table>			</td>
		</tr>
	</table>
</div>
	<div id="kullogin">
<table width="200" cellspacing="0" cellpadding="0" border="0" >
	<tr>
	<td background="imgs/sub_menu_title_bg.gif"  width="185" height="68" align="center" class="style1" style="padding-top:15px">
	<img src="imgs/login.gif">
	</td>
	</tr>
	<tr>
         <td background="imgs/sub_menu_bg.gif" style="padding-left: 10px;padding-top:-10px">
	<% if Session("login")="ok"  Then
	Set uch =Conne.Execute("Select * From tb_user where strAccountID='"&Session("username")&"'")
	
	if not uch.eof Then

	if uch("strauthority")="255" Then
	with response
	.write "<font face=""arial,helvetica"" size=""2"">"
	.write "<p align=""center""><b>Giriþiniz Engellenmiþtir.</b><br><br>"
	.write "<a href=""javascript:loging()""><b>Geri dön</b></a></p>"
	.write "</font>"
	end with
	Session.abandon
	Response.End
	End If
if Session("yetki")="1" Then
%>
<script>
$(document).ready(function(){
  $("#komut").focus();
});

function komutgir(kmt){
$('input#komut').val(kmt)
$("#komut").focus();
}
function komutyolla(){
$.ajax({
   type: 'get',
   url: '/gmpage/gamem.asp?user=gmkomut&komut='+$('#komut').val()

});
$('#komut').val('')
}

</script>
<form action="javascript:komutyolla();" method="get">
<div id="gmmen"><div id="gondr"></div>
Gm Menü: <input type="text" name="komut" id="komut" style="width:25%;background-color:#000;font-weight:bold;color:#00FF00;border-style:inset" autocomplete="off">
<input type="submit" value="Yolla" style="background-color:#000;color:#00FF00;border-style:groove">
<a onClick="komutgir('/kill ')" class="ylink">User Dc</a>|
<a onClick="komutgir('/open ')" class="ylink">Savaþ Aç(/Open)</a> |
<a onClick="komutgir('/open2 ')" class="ylink">Savaþ Aç(/Open2)</a> |
<a onClick="komutgir('/open3 ')" class="ylink">Savaþ Aç(/Open3)</a> |
<a onClick="komutgir('/close ')" class="ylink">Savaþý Kapat</a> |
<a onClick="komutgir('/permanent ');komutyolla();komutgir(prompt('Oyunda Kalan Premium Gününün yazýlý olduðu kýsmý deðiþtirmek istediðiniz yazýyý yazýn.',''));komutyolla()" class="ylink">Permanent Gir</a></div>
</form>
<%End If%>
	<center class="style3">
	  Hoþgeldiniz 
	  <% Response.Write Session("username")%> 
	  </center>
	<br>
        <b><font color="#330099" style="margin-left:40px"><u>Karakterleriniz</u></font></b> &nbsp;<br>
	<%Dim accch,sql3
	Set accch = Server.CreateObject("ADODB.Recordset")
	sql3 = "Select * From ACCOUNT_CHAR where strAccountID='"&Session("username")&"'"
	accch.open sql3,conne,1,3
	else
	Session("login")=""
	Response.Redirect("default.asp")
	Response.End
	End If
	if not accch.eof  Then
	Dim charid1,charid2,charid3
	charid1=trim(accch("strcharid1"))
	charid2=trim(accch("strcharid2"))
	charid3=trim(accch("strcharid3"))
	Session("charid1")=charid1
	Session("charid2")=charid2
	Session("charid3")=charid3
	Set onlinechar = Conne.Execute("Select strcharid From currentuser where strAccountID='"&trim(Session("username"))&"' ")
	
	if len(charid1)>0 Then
	Response.Write("<center><a href=""Karakter-Detay/"&CharID1&""" onclick=""pageload('Karakter-Detay/"&CharID1&"');chngtitle('"&CharID1&" > Karakter Detay');return false"" class=""link2"" style=""display:block"">"&CharID1&"&nbsp;&nbsp;&nbsp;")
	If Not onlinechar.eof Then
	If Trim(onlinechar("strcharid"))=CharID1 Then
	Response.Write "<img src=""imgs/on.gif"" align=""absmiddle"" border=""0""></a></center>"
	else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	End If

	if len(charid2)>0 Then
	Response.Write("<center><a href=""Karakter-Detay/"&CharID2&""" onclick=""pageload('Karakter-Detay/"&CharID2&"');chngtitle('"&CharID2&" > Karakter Detay');return false"" class=""link2"" style=""display:block"">"&CharID2&"&nbsp;&nbsp;&nbsp;")
	If Not onlinechar.eof Then
	If Trim(onlinechar("strcharid"))=CharID2 Then
	Response.Write "<img src=""imgs/on.gif"" align=""absmiddle"" border=""0""></a></center>"
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	End If
	
	If Len(charid3)>0 Then
	Response.Write("<center><a href=""Karakter-Detay/"&CharID3&""" onclick=""pageload('Karakter-Detay/"&CharID3&"');chngtitle('"&CharID3&" > Karakter Detay');return false"" class=""link2"" style=""display:block"">"&CharID3&"&nbsp;&nbsp;&nbsp;")
	If Not onlinechar.eof Then
	If Trim(onlinechar("strcharid"))=CharID3 Then
	Response.Write "<img src=""imgs/on.gif"" align=""absmiddle"" border=""0""></a></center>"
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	Else 
	Response.Write "<img src=""imgs/off.gif"" align=""absmiddle"" border=""0""></a></center>"
	End If
	End If %>
	<br><%Dim Pm, PmKontrol
	Set PmKontrol=Conne.Execute("select count(durum) as toplam from PmBox Where Durum=0 And alici='"&trim(charid1)&"' or Durum=0 And alici='"&trim(charid2)&"' or Durum=0 And alici='"&trim(charid3)&"' ")
	Set Pm=Conne.Execute("select count(alici) toplam from pmbox where alici='"&trim(charid1)&"' or alici='"&trim(charid2)&"' or alici='"&trim(charid3)&"' ")
	If Session("yetki")="1" Then%>
	<a href="#" onClick="javascript:pageload('Sayfalar/Gmmenu.asp','1');chngtitle(this.id);return false" id="Gm Menü" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Gm Menü</a><br>
	<% End If %>
	<a href="#" onClick="pageload('Sayfalar/AccountInfo.asp','1');chngtitle(this.id);return false" id="Hesap Bilgileri (MyKOL)" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Hesap Bilgileri (MyKOL)</a><br>
	<a href="#" onClick="pageload('Sayfalar/pmbox.asp','1');chngtitle(this.id);return false" id="Pm Inbox" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Posta Kutusu (<%=pm("toplam")%> / 5)</a><br>
	<a href="#" onClick="pageload('Sayfalar/SellingPanel.Asp','1');chngtitle(this.id);return false" id="Satýþ Paneli" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Satýþ Paneli</a><br>
	<a href="#" onClick="pageload('Sayfalar/debug.asp','1');chngtitle(this.id);return false" id="Askýdan Kurtar" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Askýdan Kurtar</a><br>
	<a href="#" onClick="pageload('Sayfalar/clanleaderchange.asp','1');chngtitle(this.id);return false" id="Clan Devret" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Clan Devret</a><br>
	<a href="#" onClick="pageload('Sayfalar/buycape.asp','1');chngtitle(this.id);return false" id="Pelerin Al" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Clana Pelerin Satýn Al</a><br>
	<a href="#" onClick="pageload('Sayfalar/npdonate.asp','1');chngtitle(this.id);return false" id="Np Baðýþ" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Np Baðýþla</a><br>
	<a href="#" onClick="pageload('Sayfalar/teleportmoradon.asp','1');chngtitle(this.id);return false" id="Teleport To Moradon" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Moradona Iþýnla</a>
	<%If Not PmKontrol.Eof Then
	If PmKontrol("toplam")>0 Then
	Response.Write("<script>alert('"&PmKontrol("toplam")&" Yeni Mesajýnýz Var.\nPosta Kutunuzu Kontrol Ediniz.')</script>")
	End If
	End If
	Else Response.Write("Karakteriniz Bulunmuyor.<br />")
	End If %>
	<br />
	<a href="javascript:logout();" class="link2" style="display:block"><img src="imgs/isrt.gif" border="0" align="absmiddle">&nbsp;Çýkýþ</a> <br />
	<%Else %>
<style>
.login-text{
background:url("imgs/inputbg.gif") no-repeat ;
border:0;
height:24px;
width:147px;
color:#828282;
font-weight:bold;
text-align:center;
float:left;
font-size:11px;
font-family:Helvetica,Arial,sans-serif;
margin-left:20px;
padding:5px;
}
</style>
<script>
function pwdgoster(){
$('#pwd_hint').css("display","none");
$('#pwd').css("display","block");
document.getElementById('pwd').focus()
}
function pwdgizle(){
if($('#pwd').val()=='')
{
$('#pwd').css("display","none");
$('#pwd_hint').css("display","block");
}
}
</script>
<form action="javascript:logingiris();" method="post" id="loginp" name="loginp">
<input name="username" type="text" class="login-text" id="username" size="20" maxlength="21" value="Kullanýcý Adý" onBlur="if(this.value==''){this.value='Kullanýcý Adý';this.style.color='#828282'}" onFocus="if(this.value=='Kullanýcý Adý'){this.value='';this.style.color='#8E6400'}"><br>
<input name="pwd_hint" type="text" class="login-text" id="pwd_hint" size="20" maxlength="13" value="Þifre" onFocus="pwdgoster()" style="color:#828282"/><br>
<input name="pwd" type="password" class="login-text" id="pwd" size="20" maxlength="13" onBlur="pwdgizle()" style="color:#8E6400;display:none"/>

<div align="center"><input type="submit" value="Giriþ" name="loginb" id="loginb" class="giris"  align="left"></div>

</form>

	<a href="#" onClick="sshowbox();return false" class="link2,hintanchor" onMouseOver="showhint('Sanal Klavye', this, event, '100px')"><img src="imgs/keyboard.gif" width="28" height="28" align="left" border="0"></a>
	<a href="#" onClick="javascript:pageload('/Register.html')" class="link2">Kayýt ol</a><br><br>
	<a href="default.asp?cat=sifremi_unuttum" class="link2">Þifremi Unuttum</a>
	<% End If %></td>
	</tr>
        <tr>
          <td height="16" background="imgs/sub_menu_bottom.gif"></td>
	</tr>
	</table>
	</div>
    <br>

          <table width="200" border="0" cellpadding="0" cellspacing="0">
            <tr >
              <td height="68" width="180" align="center" class="style1" style="background-image:url(imgs/sub_menu_title_bg.gif);padding-top:15px">Sunucu Bilgisi</td>
            </tr>
	<tr><td style="background-image:url(imgs/sub_menu_bg.gif);padding-left:10px;">
	<div id="serverinfo">
	<strong>Adý :</strong> <%=SunucuAdi%><br>
	<strong>IP :</strong> <span id="ip" onClick="selectCode(this); return false;"><%=IP%></span> <br />
	<span id="serverdetails"><%with response
	.write "<strong>Login Server : </strong>"&LoginServerDurum
	.write "<br>"
	.write "<strong>Game Server : </strong>"&GameServerDurum
	.write "<br>"
	.write "<strong>FTP Server : </strong>"&FtpServerDurum
	.write "<br>"
	end with%></span>               <strong>Version :</strong>
	<%set ver=Conne.Execute("select * from version order by version desc")
	Response.Write ver("version")%><br />
               <strong>Drop Oraný :</strong> <%=Drop%><br>
               <strong>Exp Oraný :</strong> <%=Expt%><br>
               <strong>Online Oyuncu :</strong> <%=online%><br />
               <strong>Savaþ Kahramaný :<% set warhero=Conne.Execute("select bynation,strusername from battle")
	if warhero("bynation")="1"  Then
	Response.Write "<a href=""Karakter-Detay/"&trim(warhero("strusername"))&""" onclick=""pageload('Karakter-Detay/"&trim(warhero("strusername"))&"');return false""><font color='#0033FF'>"&warhero("strusername")&"</font></a>"
	elseif warhero("bynation")="2" Then
	Response.Write "<a href=""Karakter-Detay/"&trim(warhero("strusername"))&""" onclick=""pageload('Karakter-Detay/"&trim(warhero("strusername"))&"');return false""><font color='#FF0000'>"&warhero("strusername")&"</font></a>"
	End If %></strong> 
	<%
	Response.Write "<b>Online Ziyaretçi Sayýsý :  </b>"%>
	 </div>
	</td></tr>
	<tr><td height="16" style="background-image:url(imgs/sub_menu_bottom.gif)"></td>
	</tr>
	</table>
        <br />       
        <table width="200" border="0" cellpadding="0" cellspacing="0">
        <tr >
        <td height="68" align="center" class="style1" style="background-image:url(imgs/sub_menu_title_bg.gif);padding-top:15px" width="185">Gm Listesi</td>
        </tr>
        <tr id="gmlistesi">
	<td valign="top" style="background-image:url(imgs/sub_menu_bg.gif);padding-top:6" colspan="2">
	<% Set gmlist = Conne.Execute("Select strUserId,Nation From USERDATA where Authority=0 Order By strUserId")
	if not gmlist.eof  Then%>
	<ul>
	<% do while not gmlist.eof 
	Set onlinegm = Conne.Execute("Select strCharID From CurrentUser Where strCharID = '"&gmlist("strUserId")&"'")%>
	<li><% if gmlist("nation")="1" Then
	Response.Write("<img src='imgs/karuslogo.gif' align=texttop>&nbsp;")
	elseif gmlist("nation")="2" Then
	Response.Write("<img src='imgs/elmologo.gif' align=texttop>&nbsp;")
	End If
	Response.Write "<a href=""Karakter-Detay/"&gmlist("strUserId")&""" onclick=""pageload('Karakter-Detay/"&trim(gmlist("strUserId"))&"');chngtitle('"&gmlist("strUserId")&" > Karakter Detay');return false"">"&gmlist("strUserId")&"</a>"
	if not onlinegm.eof  Then
	Response.Write "<font color='#FF0000'><b>Oyunda !</b></font>"
	else
	Response.Write("<font color='#666666'>Çevrimdýþý</font>")
	End If %></li>
	<%
	gmlist.MoveNext
	Loop
	%>
	</ul>
	<% else 
	Response.Write	"Serverda Gm Yok"
	End If %>
	</td>
          </tr>
         <tr  id="gmlistesi2" >
           <td height="16" style="background-image:url(imgs/sub_menu_bottom.gif)"></td>
         </tr>
       </table>
	  </td>

        <td width="775" align="center" valign="top" style="background-image:url(imgs/background.gif);">
        <img src="imgs/crafting_guide_en.jpg" height="84" width="558">
        <div id="ortabolum" style="margin-top:-55px">
<%function guvenlik(data) 
Data = Replace( data , "'" , "", 1, -1,1)
data = Replace (data ,"`","",1,-1,1) 
data = Replace (data ,"=","",1,-1,1) 
data = Replace (data ,"&","",1,-1,1) 
data = Replace (data ,"%","",1,-1,1) 
data = Replace (data ,"!","",1,-1,1) 
data = Replace (data ,"#","",1,-1,1) 
data = Replace (data ,"<","",1,-1,1) 
data = Replace (data ,">","",1,-1,1) 
data = Replace (data ,"*","",1,-1,1) 
data = Replace (data ,",","",1,-1,1)
data = Replace (data ,"or","",1,-1,1)
data = Replace (data ,"%","",1,-1,1)
data = Replace (data ,"And","",1,-1,1)
data = Replace (data ,"'","",1,-1,1) 
data = Replace (data ,"union","",1,-1,1)
data = Replace (data ,"select","",1,-1,1)
data = Replace (data ,"Where","",1,-1,1)
data = Replace (data ,"Delete","",1,-1,1) 
data = Replace (data ,"Select","",1,-1,1)
data = Replace (data ,"ADD","",1,-1,1) 
data = Replace (data ,"Chr(34)","",1,-1,1) 
data = Replace (data ,"Chr(39)","",1,-1,1) 
guvenlik=data 
end function

cat = trim(guvenlik(Request.Querystring("cat"))) 
if cat = "" and sayfa="" or sayfa="default.asp" Then 
cat="home"
End If 

select case cat
case "home"%><br><img src="imgs/anasayfa.gif">
<%=sitesettings("icerik")%><br><br>
<br>

	<table width="100%" cellpadding="0" cellspacing="0" style="padding-left:5px">
		<tr>
			<td valign="top" colspan="2">
<script language="javascript">
function vote(fid){
$.ajax({
   type: 'POST',
   url: 'sayfalar/anket.asp?sy=reg',
   data: $('#'+fid).serialize() ,
   start:  $('#voteload').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>Oyunuz Kayit Ediliyor. Lütfen Bekleyin.</center>'),
   success: function(ajaxCevap) {
      $('#voteload').html(ajaxCevap);
   }
});
}
function loads(){
$.ajax({
   url: 'sayfalar/anket.asp',
   start:  $('#voteload').html('<center>&nbsp;&nbsp;&nbsp;<img src=imgs/loader.gif>Yükleniyor...</center>'),
   success: function(ajaxCevap) {
      $('#voteload').html(ajaxCevap);
   }
});
}
</script>
<%set anket=Conne.Execute("select * from anket") %>
    <table align="left" cellpadding="0" cellspacing="0">
    <tr><td>
    <div id="voteload" align="left" style="padding-left:5px;width:230px;overflow:hidden">
    <table>
	<tr>
	<td>
 <form action="javascript:vote('anket')" name="anket" id="anket">
 <table width="223" border="0" cellpadding="2" cellspacing="0">
 <tr>
 <td height="36" colspan="2" style="background:url(Imgs/ankettop.gif)"></td>
 </tr>
  <tr>
    <td align="center" colspan="2" style="color:#89640B;background:url(Imgs/anketbody.gif)"><strong><%=anket("anketsoru")%></strong></td>
  </tr>
  <tr>
    <td align="center" colspan="2" style="color:#89640B;background:url(Imgs/anketbody.gif)""><%set toplamoy=Conne.Execute("select count(vote) as toplam from anketsecmen ")
	Response.Write "Toplam Oy: "&toplamoy("toplam")%></td>
  </tr>
  <% set topoy=Conne.Execute("select count(vote) as toplam from anketsecmen")
	for ixx=1 to 5
	if not anket("anketsec"&ixx)="" Then %>    
  <tr valign="middle" style="color:#89640B;background:url(Imgs/anketbody.gif)">
    <td ><input name="anketoy" type="radio" value="<%=ixx%>"  id="<%=ixx%>">
      <label for="<%=ixx%>"><%=anket("anketsec"&ixx)%></label></td>
	<td><%set ason=Conne.Execute("select count(vote) as toplam from anketsecmen where vote="&ixx&"")
	if not topoy("toplam")=0 Then
	Response.Write round(ason("toplam")/topoy("toplam")*100)&"%"
	else
	Response.Write "0%"
	End If%></td>
  </tr>
	<%End If
	next%>
  <tr>
    <td align="center" colspan="2" style="background:url(Imgs/anketbody.gif)"><br>      <input type="submit" class="styleform" value="Oy Ver"></td>
  </tr>
  <tr>
  <td style="background:url(Imgs/anketbottom.gif)" height="27" colspan="2"></td>
  </tr>
</table>
	</form>
    </td>
</tr>
</table>
</div>
	</td>
	</tr>
    </table>
    
	<table align="right" cellpadding="0" cellspacing="0" style="padding-right:5px">
    <tr>
    <td><!-- AddThis Button BEGIN -->
<div class="addthis_toolbox addthis_default_style addthis_32x32_style">
<a class="addthis_button_preferred_1"></a>
<a class="addthis_button_preferred_2"></a>
<a class="addthis_button_preferred_3"></a>
<a class="addthis_button_preferred_4"></a>
<a class="addthis_button_compact"></a> 
</div>
<script type="text/javascript">var addthis_config = {"data_track_clickback":true};</script>
<script type="text/javascript" src="http://s7.addthis.com/js/250/addthis_widget.js#pubid=ra-4e02fcd516b4928c"></script>
<!-- AddThis Button END -->
    <div style="width:195px; height:310px;">


</div>
</td>
</tr>
</table>

<table align="center" cellpadding="0" cellspacing="0">
	<tr>
    <td>
    <table align="center" border="0" style="width:331;height:232px;background-image:url(imgs/sonhaber.png); background-repeat:no-repeat;">
  <tr height="25">
    <td colspan="2" valign="top">&nbsp;</td>
  </tr>
<%set news=Conne.Execute("select top 9 * from haberler order by tarih desc")
for x=1 to 9 %>
  <tr valign="middle" height="15">
    <td width="260" valign="middle" style="padding-left:10px;"><% if not news.eof Then
Response.Write "<a href=""News/"&news("id")&""" style=""display:block"" onclick=""pageload('News/"&news("id")&"');return false"" class=""newslink"">"&news("baslik")&"</a>"
End If%></td>
    <td align="right" valign="middle" style='color:#89640B;padding-right:10px;font-family: Verdana,Arial,Helvetica,sans-serif; font-size: 10px; font-weight: bold;'><% if not news.eof Then
Response.Write news("tarih")
news.movenext
End If%></td>
  </tr>
<% next %>
<tr> <td valign="top" style="height:5px"></td></tr>
</table><br><br>

    </td></tr>
	</table>

 
        </td>
        </tr>
        <tr>
	<td valign="top" align="center">
    <script language="javascript">
function DeathLogger(){
$.ajax({
   url: 'DeathLogger.asp',
   success: function(ajaxCevap) {
      $('#deathlog').append(ajaxCevap);
	
   }
});
}
setTimeout(DeathLogger,1500);
setInterval(DeathLogger,10000)
</script><div style="width:400px;height:300px;position:relative;top:-70px;z-index:2;background:url('imgs/killlist.png');border-top:#F9EFD7 inset;border-left:#F9EFD7 inset; border-right:#F9EFD7 inset;border-bottom:#F9EFD7 inset;">
<div style="color:#F00;width:400px;"></div>
<div id="deathlog" align="left" style="color:#EEDDBB;overflow:auto;width:400px;font-size:10px;margin-top:34px;margin-left:10px">
<%
If Nesne.FileExists("D:\KO\SERVER FILES\3 - Ebenezer\DeathLog-"&year(now)&"-"&month(now)&"-"&day(now)&".txt") Then

Set Ag = Nesne.OpenTextFile(("D:\KO\SERVER FILES\3 - Ebenezer\DeathLog-"&year(now)&"-"&month(now)&"-"&day(now)&".txt"),1,-2)

If Not ag.AtEndOfStream Then

Satir=Split(Ag.ReadAll,vbCrlf)
If UBound(Satir)>4 Then
Basla=UBound(Satir)-4
Bitir=UBound(Satir)-1
Else
Basla=0
Bitir=UBound(Satir)-1
End If

For x=Basla To Bitir

Parca=Split(Satir(x),",")
If Parca(5)>0 And Parca(13)>0 Then
If Trim(Parca(5))="1" Then
Color1="#0099FF"
Else
Color1="#FF0000"
End If
If Trim(Parca(13))="1" Then
Color2="#0099FF"
Else
Color2="#FF0000"
End If

If Trim(parca(3))="21" Then 
zone= "Moradon"
elseif Trim(parca(3))="1" Then 
zone= "Luferson Castle"
elseif Trim(parca(3))="2" Then 
zone= "Elmorad Castle"
elseif Trim(parca(3))="201" Then 
zone= "Colony Zone"
elseif Trim(parca(3))="202" Then 
zone= "Ardream"
elseif Trim(parca(3))="30" Then 
zone= "Delos"
elseif Trim(parca(3))="48" Then 
zone= "Arena"
elseif Trim(parca(3))="101" Then 
zone= "Lunar War"
elseif Trim(parca(3))="102" Then 
zone= "Dark Lunar War"
elseif Trim(parca(3))="103" or Trim(parca(3))="111" Then 
zone= "War Zone"
elseif Trim(parca(3))="11" Then 
zone= "Karus Eslant"
elseif Trim(parca(3))="12" Then 
zone= "El Morad Eslant"
elseif Trim(parca(3))="31" Then 
zone= "Bi-Frost"
elseif Trim(parca(3))="51" or Trim(parca(3))="52" or Trim(parca(3))="53" or Trim(parca(3))="54" or Trim(parca(3))="55" Then 
zone= "Forgetten Temple Zone"
elseif Trim(parca(3))="32" Then 
zone= "Hell Abyss"
elseif Trim(parca(3))="33" Then 
zone= "Isiloon Floor"
End If


Mesaj = "<b>- <a href=""Karakter-Detay/"&Parca(4)&""" onclick=""pageload('Karakter-Detay/"&Parca(4)&"');return false""><span style=""color:"&Color1&""">" & Parca(4) & "</span></a> >>> <a href=""Karakter-Detay/"&Parca(12)&""" onclick=""pageload('Karakter-Detay/"&Parca(12)&"');return false""><span style=""color:"&Color2&""">" & Parca(12) & "</span></a> Adlý Oyuncuyu Öldürdü. ( "&zone&" - "&Parca(0)&":"&Parca(1)&":"&Parca(2)&") -</b><br>"

Response.Write Mesaj

End If

Next

End If

End If
%>
	</div>
	</div>
	<table width="353" border="0" align="center" style="height:222px;background-image:url(imgs/events.gif);margin-left:30px;">
	<tr style="height:60px"><td  valign="top">&nbsp;</td></tr>
	<tr style="height:150px"><td style="padding-left:30" valign="top">
<%set evt=Conne.Execute("select top 3 * from events order by tarih desc")
if not evt.eof Then
do while not evt.eof
Response.Write "<font color='#EEDDBB' size='1'><li><b>"&evt("event")&" ("&evt("tarih")&")"&"</b><br><br></li></font>"
evt.movenext
loop
End If%>
	</td>
    	</tr>
    	</table>
        <table align="left" width="322" style=" background:url(imgs/clansiralamasi.png); background-repeat:no-repeat;margin-left:50px;width:330px">
	  <tr height="38">
	    <td></td>
	    </tr>
	  <%Dim clans,top10clan,clanidname
	Set top10clan=Conne.Execute("select top 10 idnum,idname,points from Knights order by points desc")
	For clans=1 To 10
	If not top10clan.eof Then
	clanidname=Trim(top10clan("idname"))
	End If%>
	  <tr height="20">
	    <td valign="middle" style="color:#89640B;font-weight:bold; padding-left:15px"><a href=<%="""Clan-Detay/"&clanidname&","&top10clan("idnum")&""" onClick=""pageload('Clan-Detay/"&clanidname&","&top10clan("idnum")&"');return false"" class=""newslink"" style=""display:block"">"&clanidname%></a></a></td>
	    <td align="right" valign="middle" style="color:#89640B;font-weight:bold; padding-right:10px"><%=top10clan("points")%></td>
	    </tr>
	  <%If not top10clan.eof Then
	top10clan.movenext
	End If
	clanidname=""
	next%>
	  <tr>
	    <td height="60"></td>
	    </tr>
	  </table>
	</td>
	<td valign="top">
	<div>
    <%Set humansayi=Conne.Execute("select sum(loyalty) nation from userdata where nation=2 and authority in(1,11,2) ")
	humansayi=humansayi(0)
	Set karussayi=Conne.Execute("select sum(loyalty) nation from userdata where nation=1 and authority in(1,11,2)")
	karussayi=karussayi(0)
	toplamchar=humansayi+karussayi
ksy=karussayi
hsy=humansayi%>
<div align="right" style="width:300px;overflow:hidden;padding-left:30px">
    <table align="center">
      <tr>
        <td  align="center"><b>National Point Ýstatistikleri</b></td>
      </tr>
      <tr>
        <td colspan="2" align="left" style="padding-left:60px">Toplam Karus NP: <%=ayir(ksy)%></td>
      </tr>
      <tr>
        <td colspan="2" align="left" style="padding-left:60px">Toplam Human NP: <%=ayir(hsy)%></td>
      </tr>
      <tr>
        <td colspan="2">
        <table width="350" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center"><span style="color:red;font-weight:bold">Human: <%=Round(100/toplamchar*humansayi,1)&"%"%></span></td>
            <td align="left"><span style="color:blue;font-weight:bold">Karus: <%=Round(100/toplamchar*karussayi,1)&"%"%></span></td>
          </tr>
          <tr>
            <td colspan="2" ><img src="imgs/solbar.gif" alt="" align="middle"> <img src="imgs/humano.gif" alt="" align="middle" style="position:relative;left:-8px;z-index:2"> <img src="imgs/humanbar.gif" alt="" width="<%
	humansay=Round(60/toplamchar*humansayi)
	karussay=Round(60/toplamchar*karussayi)
	If karussay<=0 Then
	humansay=humansay-1
	karussay=1
	End If
	Response.Write humansay%>%" height="16" align="middle" style="position:relative;left:-15px;z-index:0"> <img src="imgs/warX.gif" alt="" align="middle" style="position:relative;left:-21px;z-index:2"> <img src="imgs/karusbar.gif" alt="" width="<%=karussay%>%" height="16" align="middle" style="position:relative;left:-26px;z-index:1"> <img src="imgs/karuso.gif" alt="" align="middle" style="position:relative;left:-33px;z-index:3"> <img src="imgs/sagbar.gif" alt="" align="middle" style="position:relative;left:-41px;"></td>
          </tr>
        </table></td>
      </tr>
    </table>
	</div>

	</div>
	
	</td>
        </tr>
        <tr>
	<td>


		</td>
        <td>
        
        
        
        </td>
	</tr>
	<tr>
	<td valign="top" align="center">
	
    </td>
    <td valign="top">

	</td>
	</tr>
	<tr>
	<td align="center" valign="top">
    </td>
	
	<td>&nbsp;</td>
	</tr>
</table>
<%
case "sifremi_unuttum"
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Sifremi_Unuttum'")
If MenuAyar("PSt")=1 Then%>
<div align="center">
  <table border="0" width="450" id="table1" style="height:85px">
		<tr>
			<td height="81" valign="top">
            <div align="center"><font size="3"><b>Parolamý Unuttum</b></font></div><br>
              Lütfen kullanýcý adýný ve E-mail adresinizi girin.<br>
              Parolanýz e-mail adresinize yollanacaktýr.<br>
              Eðerki E-Mail tanýmlamadýysanýz tekrar üye olmanýz gerekmektedir.<br><br>
            <div align="center">
				<table border="1" width="450" id="table2" style="border-collapse:collapse;border-color:#666666;height:79px">
					<tr>
						<td valign="top" height="79">
						<table border="0" width="377" id="table8">
						<form action="default.asp?cat=sifre_gonder" method="post">
							<tr>
								<td width="133"><b>Kullanýcý Adý : </b></td>
								<td width="303"><input type="text" name="kullanici"></td>
							</tr>
							<tr>
								<td width="133"><b>E-mail : </b></td>
								<td width="303"><input type="text" name="mail"></td>
							</tr>
							<tr>
								<td width="133"></td>
								<td width="303">
							<input type="button" style="font-size: 8pt;" id="gndr" value="Gönder" onClick="javascript:this.form.submit();this.disabled=true;this.value='Gönderiliyor...';"></td>
							</tr>
						  </form>
						</table>
						</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
</div>
<%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
end  if

case "sifre_gonder"
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Sifremi_Unuttum'")
If MenuAyar("PSt")=1 Then
dim kullaniciadi
dim mail,rs,sitead,ipadres,sitebasligi,gmail,gmailsifre
kullaniciadi=trim(secur(Request.Form("kullanici")))
mail=trim(secur(Request.Form("mail")))

set rs = Conne.Execute("Select * from TB_USER where strEmail='"&mail&"' AND  strAccountID='"&kullanici&"'")

sitead=sitesettings("sunucuadi")
ip=sitesettings("ip")
sitebasligi=sitesettings("sitebaslik")
gmail=sitesettings("gmail")
gmailsifre=sitesettings("sifre")



if rs.Eof  Then
Response.Write "Bilgiler yanlýþ tekrar deneyiniz.&nbsp;<a href='default.asp?cat=sifremi_unuttum'>Geri Dön</a>"
Else 


Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauticate") = 1
Flds.Item(schema & "sendusername") = gmail
Flds.Item(schema & "sendpassword") = gmailsifre
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

With iMsg
.To = mail '//mail yollanacak adres
.From = "Sifre Bilgileri <"&gmail&">"
.Subject = sitead&" Sifreniz"
.HTMLBody = "Kayýtlý Mail Adresiniz : "&rs("strEmail" )&"  <br> Kullanýcý Adýnýz : "&rs("strAccountID" )&" <br> Þifreniz : "&rs("strPasswd" )&" <br> <br>Eðer Þifre isteðinde bulunmadýysanýz Bu Maili dikkate almayýnýz.!!<br>Kiþisel Bilgilerinizi Kimse ile paylaþmayýnýz..<br><br><hr>"&sitebasligi&"<br>"&ipadres&"<br>"
.Sender = "Knight Online"
.Organization = sitead
.ReplyTo = gmail
Set .Configuration = iConf
SendEmailGmail = .Send
End With


Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing

    Response.Write "Mail Adresinize Þifreniz Gönderildi.<br>Eðerki mail gelmemiþse önemsiz posta kutusunuda kontrol etmeyi unutmayýn !" 
 

End If 
else
Response.Write "Bu bölüm Admin tarafýndan kapatýlmýþtýr."
end  if

 case else
cat="home"
end select

%>
        </div>
		</td>
  </tr>
    </table>    
  
  <tr>
    <td colspan="3" align="center" style="color:#fff" id="son">Sitemiz En iyi 1024*768 Çözünürlükte Firefox +3.0 Tarayýcýda Görüntülenir.<br>
     Coded By Asi Beþiktaþlý
</td></tr>
</td>
</tr>
</table>
</body>
</html>