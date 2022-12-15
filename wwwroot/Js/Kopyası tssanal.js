/* Virtual Keyboard added By SyMurG, http://www.turksecurity.org, info@turksecurity.org */


	// source is the field which is currently focused:
   var source = null;

   function search_for_input_field(num)
   {
      var tg = document.getElementsByTagName("INPUT");
      var tgt = document.getElementsByTagName("TEXTAREA");

      if(tg && tg[num])
        return tg[num];

      if(tgt && tgt[num])
        return tgt[num];

      return null;
   }

   // This function retrieves the source element
   // for the given event object:
   function get_event_source(e)
   {
     var event = e ? e : window.event;
     return event.srcElement ? event.srcElement : event.target;
   }

   // This function binds handler function to the 
   // eventType event of the elem element:
   function setup_event(elem, eventType, handler)
   {
     return (elem.attachEvent) ? elem.attachEvent("on" + eventType, handler) : ((elem.addEventListener) ? elem.addEventListener(eventType, handler, false) : false);
   }

   // By focusing the INPUT field we set the source
   // to the newly focused field:
   function focus_keyboard(e)
   {
     source = get_event_source(e);
   }

   // This function slightly differs from one with the same name
   // in 4-test-fly sample. Now it accepts not the id, but the
   // number (index in the INPUT elements array) of the INPUT field.
   function register_field(num)
   {
     var tg = document.getElementsByTagName("INPUT");
     var tgt = document.getElementsByTagName("TEXTAREA");

     if(tg && tg[num])
       setup_event(tg[num], "focus", focus_keyboard);

     if(tgt && tgt[num])
       setup_event(tgt[num], "focus", focus_keyboard);
   }

   // This function enumerates and "registers" all INPUT fields
   // on the page:
   function register_input_fields()
   {
     var tg = document.getElementsByTagName("INPUT");
     var tgt = document.getElementsByTagName("TEXTAREA");

     if(tg)
     {
       for(var i = 0; i < tg.length; i++)
         register_field(i);
     }
     if(tgt)
     {
       for(var i = 0; i < tgt.length; i++)
         register_field(i);
     }
   }

   function keyb_callback(ch)
   {
     switch(ch)
     {
       case "BackSpace":
         var min = (source.value.charCodeAt(source.value.length - 1) == 10) ? 2 : 1;
         source.value = source.value.substr(0, source.value.length - min);
         break;

       case "Enter":
         source.value += "\n";
         break;

       default:
         source.value += ch;
     }

     source.focus();
   }

   function sinit()
   {
     // Note: all parameters, starting with 3rd, in the following
     // expression are equal to the default parameters for the
     // VKeyboard object. The only exception is 15th parameter
     // (flash switch), which is false by default.

     new VKeyboard("keyboard",    // containers id
                   keyb_callback, // reference to the callback function
                   true,          // create the numpad or not? (this and the following params are optional)
                   "",            // font name ("" == system default)
                   "12px",        // font size in px
                   "#000",        // font color
                   "#F00",        // font color for the dead keys
                   "#FFF",        // keyboard base background color
                   "#FFF",        // keys' background color
                   "#DDD",        // background color of switched/selected item
                   "#777",        // border color
                   "#CCC",        // border/font color of "inactive" key (key with no value/disabled)
                   "#FFF",        // background color of "inactive" key (key with no value/disabled)
                   "#F77",        // border color of the language selector's cell
                   true,          // show key flash on click? (false by default)
                   "#CC3300",     // font color for flash event
                   "#FF9966",     // key background color for flash event
                   "#CC3300",     // key border color for flash event
                   true);        // embed VKeyboard into the page?

     // The very 1st (index == 0) field is "focused" by default:
     source = search_for_input_field(0);

     // Any INPUTs? Register them all!
     if(source) register_input_fields();
   }

var sns4=document.layers
var sie4=document.all
var sns6=document.getElementById&&!document.all

//drag drop function for NS 4////
/////////////////////////////////

var sdragswitch=0
var snsx
var snsy
var snstemp

function sdrag_dropns(name){
if (!sns4)
return
temp=eval(name)
temp.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP)
temp.onmousedown=sgons
temp.onmousemove=sdragns
temp.onmouseup=sstopns
}

function sgons(e){
temp.captureEvents(Event.MOUSEMOVE)
snsx=e.x
snsy=e.y
}
function sdragns(e){
if (sdragswitch==1){
temp.moveBy(e.x-snsx,e.y-snsy)
return false
}
}

function sstopns(){
temp.releaseEvents(Event.MOUSEMOVE)
}

//drag drop function for ie4+ and NS6////
/////////////////////////////////


function sdrag_drop(e){
if (sie4&&dragapproved){
crossobj.style.left=tempx+event.clientX-offsetx
crossobj.style.top=tempy+event.clientY-offsety
return false
}
else if (sns6&&dragapproved){
crossobj.style.left=tempx+e.clientX-offsetx+"px"
crossobj.style.top=tempy+e.clientY-offsety+"px"
return false
}
}

function sinitializedrag(e){
crossobj=sns6? document.getElementById("sshowimage") : document.all.sshowimage
var firedobj=sns6? e.target : event.srcElement
var topelement=sns6? "html" : document.compatMode && document.compatMode!="BackCompat"? "documentElement" : "body"
while (firedobj.tagName!=topelement.toUpperCase() && firedobj.id!="sdragbar"){
firedobj=sns6? firedobj.parentNode : firedobj.parentElement
}

if (firedobj.id=="sdragbar"){
offsetx=sie4? event.clientX : e.clientX
offsety=sie4? event.clientY : e.clientY

tempx=parseInt(crossobj.style.left)
tempy=parseInt(crossobj.style.top)

dragapproved=true
document.onmousemove=sdrag_drop
}
}
document.onmouseup=new Function("dragapproved=false")

////drag drop functions end here//////

function shidebox(){
crossobj=sns6? document.getElementById("sshowimage") : document.all.sshowimage
crossobj2=sns6? document.getElementById("sytopbar") : document.all.sytopbar
if (sie4||sns6)
crossobj.style.visibility="hidden"
else if (sns4)
document.sshowimage.visibility="hide"
}

function sshowbox(){
crossobj=sns6? document.getElementById("sshowimage") : document.all.sshowimage
crossobj2=sns6? document.getElementById("sytopbar") : document.all.sytopbar
if (sie4||sns6)
crossobj.style.visibility="visible"
else if (sns4)
document.sshowimage.visibility="visible"
displaysshowimage()
}

function displaysshowimage(){
var ie=document.all && !window.opera
var dom=document.getElementById
iebody=(document.compatMode=="CSS1Compat")? document.documentElement : document.body
objref=(dom)? document.getElementById("sshowimage") : document.all.sshowimage
var scroll_top=(ie)? iebody.scrollTop : window.pageYOffset
var docwidth=(ie)? iebody.clientWidth : window.innerWidth
docheight=(ie)? iebody.clientHeight: window.innerHeight
var objwidth=objref.offsetWidth
objheight=objref.offsetHeight
objref.style.left=docwidth/2-objwidth/2+"px"
objref.style.top=scroll_top+docheight/2-objheight/2+"px"
}

if (window.addEventListener)
window.addEventListener("load", sinit, false)
else if (window.attachEvent)
window.attachEvent("onload", sinit)
else if (document.getElementById)
window.onload=sinit