/* Virtual Keyboard added By SyMurG, http://www.turksecurity.org, info@turksecurity.org */

//Specify the text to display
var displayed='<font color="#ff0000">TS-Sanal Klavye</font>'

///////////////////////////Do not edit below this line////////////

var logolink='#TurkSecurity.Org'
var logoclick= 'sshowbox();return false'
var ns4=document.layers
var ie4=document.all
var ns6=document.getElementById&&!document.all

function ietruebody(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function regenerate(){
window.location.reload()
}
function regenerate2(){
if (ns4)
setTimeout("window.onresize=regenerate",400)
}                                            



function staticit(){ //function for IE4/ NS6
var w2=ns6? pageXOffset+w : ietruebody().scrollLeft+w
var h2=ns6? pageYOffset+h : ietruebody().scrollTop+h
crosslogo.style.left=w2+"px"
crosslogo.style.top=h2+"px"
}

function staticit2(){ //function for NS4
staticimage.left=pageXOffset+window.innerWidth-staticimage.document.width-28
staticimage.top=pageYOffset+window.innerHeight-staticimage.document.height-10
}



