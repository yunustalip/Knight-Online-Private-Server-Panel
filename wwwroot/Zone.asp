<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<%response.expires=0
zoneid=secur(Request.Querystring("zoneid"))
Dim MenuAyar,ksira,zoneid,map,px,pz,cleft,ctop,users,x
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='Kim-Nerede'")
If MenuAyar("PSt")=1 Then

if isnumeric(zoneid)=true Then
zoneid=cint(zoneid)
else
zoneid="21"
End If

map=zoneid

if zoneid="1" or zoneid="2" or zoneid="21" or zoneid="201" or zoneid="202" or zoneid="11" or zoneid="12" Then

if zoneid=11 or zoneid=12 Then
map=1112
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
set users=Conne.Execute("select u.struserid, u.level, u.zone, u.px, u.pz, u.loyalty, u.loyaltymonthly from userdata u,currentuser c where c.strCharID=u.strUserID and u.zone='"&zoneid&"'")
x=60
if not users.eof Then
do while not users.eof
if len(users("px"))=6 Then
px=left(users("px"),4)
elseif len(users("px"))=5 Then
px=left(users("px"),3)
End If
if len(users("pz"))=6 Then
pz=left(users("pz"),4)
elseif len(users("pz"))=5 Then
pz=left(users("pz"),3)
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
Response.Write "<div name="""&trim(users("struserid"))&""" style=""position:absolute;left:"&cleft&"px; top:"&ctop&"px; "" id="""&trim(users("struserid"))&"""><img id="""&users("struserid")&""" src=""imgs/nokta.gif"" onmouseover=""showhint('"&trim(users("struserid"))&"("&users("level")&")<br>"&users("loyalty")&" / "&users("loyaltymonthly")&"',this, event,'120px')""></div>"&vbcrlf
Response.Write "<div style=""position:absolute;left:520;top:"&x&""" onMouseOut=""karakterkucult('"&users("struserid")&"')"" onMouseOver=""karakterbul('"&users("struserid")&"');showhint('"&trim(users("struserid"))&"("&users("level")&")<br>"&users("loyalty")&" / "&users("loyaltymonthly")&"',this, event,'120px');""><a >"&users("struserid")&"</a></div>"&vbcrlf
x=x+15
users.movenext
loop
End If
else
Response.Write "Bu Haritada Kimse Bulunmamaktadýr."
End If
%>
<div style="position:absolute;left:520;top:8"><form action="" method="get">
Harita: <select name="zoneid" onChange="this.form.submit()" class="styleform">
<option value="21" <%if zoneid="21" Then Response.Write "selected"%>>Moradon</option>
<option value="1" <%if zoneid="1" Then Response.Write "selected"%>>Luferson</option>
<option value="2" <%if zoneid="2" Then Response.Write "selected"%>>El Morad Castle</option>
<option value="201" <%if zoneid="201" Then Response.Write "selected"%>>Colony Zone</option>
<option value="202" <%if zoneid="202" Then Response.Write "selected"%>>Ardream</option>
<option value="11" <%if zoneid="11" or zoneid="12" Then Response.Write "selected"%>>Eslant</option>
</select>
<br><br><center>Online Oyuncular<hr></center>
</form></div>
<%MenuAyar.Close
Set MenuAyar=Nothing

Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If%>