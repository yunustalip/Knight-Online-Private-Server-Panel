<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%response.expires=0
zoneid=secur(Request.Querystring("zoneid"))
pxx=secur(Request.Querystring("px"))
pzz=secur(Request.Querystring("pz"))

if isnumeric(zoneid)=false or isnumeric(px)=false or isnumeric(pz)=false Then
Response.End
End If

if zoneid="1" or zoneid="2" or zoneid="21" or zoneid="201" or zoneid="202" or zoneid="11" or zoneid="12" Then

if zoneid=11 or zoneid=12 Then
map=1112
else
map=zoneid
End If%>

<style type="text/css">
body {
background-image:url(imgs/Maps/<%Response.Write map&".jpg"%>);
background-position:top left;
background-repeat:no-repeat;
margin-left:0px;
margin-top:0px;

}

#hintbox{ /*CSS for pop up hint box */
position:absolute;
top: 0;
background-color: lightyellow;
width: 150px; /*Default width of hint.*/ 
padding: 3px;
border:1px solid black;
font:normal 11px Verdana;
line-height:18px;
z-index:100;
border-right: 3px solid black;
border-bottom: 3px solid black;
visibility: hidden;
}

.hintanchor{ /*CSS for link that shows hint onmouseover*/
font-weight: bold;
color: navy;
margin: 3px 8px;
}

</style>

<script type="text/javascript">


var horizontal_offset="9px" //horizontal offset of hint box from anchor link

/////No further editting needed

var vertical_offset="0" //horizontal offset of hint box from anchor link. No need to change.
var ie=document.all
var ns6=document.getElementById&&!document.all

function getposOffset(what, offsettype){
var totaloffset=(offsettype=="left")? what.offsetLeft : what.offsetTop;
var parentEl=what.offsetParent;
while (parentEl!=null){
totaloffset=(offsettype=="left")? totaloffset+parentEl.offsetLeft : totaloffset+parentEl.offsetTop;
parentEl=parentEl.offsetParent;
}
return totaloffset;
}

function iecompattest(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function clearbrowseredge(obj, whichedge){
var edgeoffset=(whichedge=="rightedge")? parseInt(horizontal_offset)*-1 : parseInt(vertical_offset)*-1
if (whichedge=="rightedge"){
var windowedge=ie && !window.opera? iecompattest().scrollLeft+iecompattest().clientWidth-30 : window.pageXOffset+window.innerWidth-40
dropmenuobj.contentmeasure=dropmenuobj.offsetWidth
if (windowedge-dropmenuobj.x < dropmenuobj.contentmeasure)
edgeoffset=dropmenuobj.contentmeasure+obj.offsetWidth+parseInt(horizontal_offset)
}
else{
var windowedge=ie && !window.opera? iecompattest().scrollTop+iecompattest().clientHeight-15 : window.pageYOffset+window.innerHeight-18
dropmenuobj.contentmeasure=dropmenuobj.offsetHeight
if (windowedge-dropmenuobj.y < dropmenuobj.contentmeasure)
edgeoffset=dropmenuobj.contentmeasure-obj.offsetHeight
}
return edgeoffset
}

function showhint(menucontents, obj, e, tipwidth){
if ((ie||ns6) && document.getElementById("hintbox")){
dropmenuobj=document.getElementById("hintbox")
dropmenuobj.innerHTML=menucontents
dropmenuobj.style.left=dropmenuobj.style.top=-500
if (tipwidth!=""){
dropmenuobj.widthobj=dropmenuobj.style
dropmenuobj.widthobj.width=tipwidth
}
dropmenuobj.x=getposOffset(obj, "left")
dropmenuobj.y=getposOffset(obj, "top")
dropmenuobj.style.left=dropmenuobj.x-clearbrowseredge(obj, "rightedge")+obj.offsetWidth+"px"
dropmenuobj.style.top=dropmenuobj.y-clearbrowseredge(obj, "bottomedge")+"px"
dropmenuobj.style.visibility="visible"

}
}

function hidetip(e){
dropmenuobj.style.visibility="hidden"
dropmenuobj.style.left="-500px"
}

function createhintbox(){
var divblock=document.createElement("div")
divblock.setAttribute("id", "hintbox")
document.body.appendChild(divblock)
}

if (window.addEventListener)
window.addEventListener("load", createhintbox, false)
else if (window.attachEvent)
window.attachEvent("onload", createhintbox)
else if (document.getElementById)
window.onload=createhintbox

</script>
<style>
body{
	color: #808080;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
	text-decoration: none;
	font-weight: bold;

}
.sar{
	color: #FFFFFF;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:12px;
	text-decoration: none;
	font-weight: bold;
}
.styleform {
	color: #000000;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}

</style>

<script>
function karakterbul(karakter){
document.getElementById(karakter).src='imgs/noktak.gif';
}
function karakterkucult(karakter){
document.getElementById(karakter).src='imgs/nokta.gif';
}

</script>
<body bgcolor="#F9EED8">

<%
x=60

if len(pxx)=6 Then
px=left(pxx,4)
elseif len(pxx)=5 Then
px=left(pxx,3)
End If
if len(pzz)=6 Then
pz=left(pzz,4)
elseif len(pzz)=5 Then
pz=left(pzz,3)
End If

if zoneid="1" or zoneid="2" or zoneid="201" Then
px=round(px/4)
pz=round(pz/4)
End If
if zoneid="11" or zoneid="12" or zoneid="202" Then
px=round(px/2)
pz=round(pz/2)
End If

if  zoneid="21" and px>511 or pz>511 Then
px=306
pz=352
End If
cleft=px
ctop=511-pz
Response.Write "<div name="""" style=""position:relative;left:"&cleft&"px; top:"&ctop-20&"px; "" id=""""><img id=""x"" src=""imgs/Red_Arrow_Down.gif"" width=""15"" height=""20"" onmouseover=""showhint('')<br> / ',this, event,'120px')""></div>"&vbcrlf
x=x+15


else
Response.Write "Bu Haritada Bulunmamaktad�r."
End If
%>
