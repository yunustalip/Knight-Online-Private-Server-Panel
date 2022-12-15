<!--#include file="../_inc/conn.asp"-->
<script src="../js/jquery.js" type="text/javascript"></script>
<style>body{color:#000;font-size:12px}</style>
<%zoneid=Request.Querystring("zoneid")
if zoneid="" Then Zoneid="21"
set npcs=Conne.Execute("select knp.zoneid,knp.npcid,knp.id,k.strname,k.bygroup,k.ssid,knp.leftx,knp.topz from k_npcpos knp, k_npc k where knp.zoneid="&zoneid&" and knp.npcid=k.ssid and acttype<>1 group by knp.zoneid,knp.npcid,k.strname,k.bygroup,k.ssid,knp.id,knp.leftx,knp.topz")
If Npcs.Eof Then
Response.Write "Npc Bulunamadý."
Response.End 
End If

Dim NpcZone
Set NpcZone=Conne.Execute("Select bz From Zone_Info Where ZoneNo="&zoneid)%>
<script type="text/javascript">
var Drag = {

     obj : null,

     init : function(o, oRoot, minX, maxX, minY, maxY, bSwapHorzRef, bSwapVertRef, fXMapper, fYMapper)
     {
          o.onmousedown     = Drag.start;

          o.hmode               = bSwapHorzRef ? false : true ;
          o.vmode               = bSwapVertRef ? false : true ;

          o.root = oRoot && oRoot != null ? oRoot : o ;

          if (o.hmode  && isNaN(parseInt(o.root.style.left  ))) o.root.style.left   = "0px";
          if (o.vmode  && isNaN(parseInt(o.root.style.top   ))) o.root.style.top    = "0px";
          if (!o.hmode && isNaN(parseInt(o.root.style.right ))) o.root.style.right  = "0px";
          if (!o.vmode && isNaN(parseInt(o.root.style.bottom))) o.root.style.bottom = "0px";

          o.minX     = typeof minX != 'undefined' ? minX : null;
          o.minY     = typeof minY != 'undefined' ? minY : null;
          o.maxX     = typeof maxX != 'undefined' ? maxX : null;
          o.maxY     = typeof maxY != 'undefined' ? maxY : null;

          o.xMapper = fXMapper ? fXMapper : null;
          o.yMapper = fYMapper ? fYMapper : null;

          o.root.onDragStart     = new Function();
          o.root.onDragEnd     = new Function();
          o.root.onDrag          = new Function();
     },

     start : function(e)
     {
          var o = Drag.obj = this;
          e = Drag.fixE(e);
          var y = parseInt(o.vmode ? o.root.style.top  : o.root.style.bottom);
          var x = parseInt(o.hmode ? o.root.style.left : o.root.style.right );
          o.root.onDragStart(x, y);

          o.lastMouseX     = e.clientX;
          o.lastMouseY     = e.clientY;

          if (o.hmode) {
               if (o.minX != null)     o.minMouseX     = e.clientX - x + o.minX;
               if (o.maxX != null)     o.maxMouseX     = o.minMouseX + o.maxX - o.minX;
          } else {
               if (o.minX != null) o.maxMouseX = -o.minX + e.clientX + x;
               if (o.maxX != null) o.minMouseX = -o.maxX + e.clientX + x;
          }

          if (o.vmode) {
               if (o.minY != null)     o.minMouseY     = e.clientY - y + o.minY;
               if (o.maxY != null)     o.maxMouseY     = o.minMouseY + o.maxY - o.minY;
          } else {
               if (o.minY != null) o.maxMouseY = -o.minY + e.clientY + y;
               if (o.maxY != null) o.minMouseY = -o.maxY + e.clientY + y;
          }

          document.onmousemove     = Drag.drag;
          document.onmouseup          = Drag.end;

          return false;
     },

     drag : function(e)
     {

          e = Drag.fixE(e);
          var o = Drag.obj;

          var ey     = e.clientY;
          var ex     = e.clientX;
          var y = parseInt(o.vmode ? o.root.style.top  : o.root.style.bottom);
          var x = parseInt(o.hmode ? o.root.style.left : o.root.style.right );
          var nx, ny;

          if (o.minX != null) ex = o.hmode ? Math.max(ex, o.minMouseX) : Math.min(ex, o.maxMouseX);
          if (o.maxX != null) ex = o.hmode ? Math.min(ex, o.maxMouseX) : Math.max(ex, o.minMouseX);
          if (o.minY != null) ey = o.vmode ? Math.max(ey, o.minMouseY) : Math.min(ey, o.maxMouseY);
          if (o.maxY != null) ey = o.vmode ? Math.min(ey, o.maxMouseY) : Math.max(ey, o.minMouseY);

          nx = x + ((ex - o.lastMouseX) * (o.hmode ? 1 : -1));
          ny = y + ((ey - o.lastMouseY) * (o.vmode ? 1 : -1));

          if (o.xMapper)          nx = o.xMapper(y)
          else if (o.yMapper)     ny = o.yMapper(x)

          Drag.obj.root.style[o.hmode ? "left" : "right"] = nx + "px";
          Drag.obj.root.style[o.vmode ? "top" : "bottom"] = ny + "px";
          Drag.obj.lastMouseX     = ex;
          Drag.obj.lastMouseY     = ey;
		  var keyprBil = document.getElementById('keypressBilgi');
	<%if zoneid="1" or zoneid="2" or zoneid="201" Then
          Response.Write("document.getElementById('dpx').value=((nx+3)*2);"&vbcrlf)
          Response.Write("document.getElementById('dpy').value=((1023-ny-15)*2);"&vbcrlf)
          Response.Write("document.getElementById('did').value=o.id;"&vbcrlf)
		  Response.Write("keyprBil.innerHTML = 'PosX: '+((nx+3)*2)+ ' PosY: ' + ((1023-ny-15)*2);"&vbcrlf)
	Elseif zoneid="11" or zoneid="12" or zoneid="202" Then
	Response.Write("document.getElementById('dpx').value=((nx+3));"&vbcrlf)
    Response.Write("document.getElementById('dpy').value=((1023-ny-15));"&vbcrlf)
    Response.Write("document.getElementById('did').value=o.id;"&vbcrlf)
	Elseif zoneid="21" Then
	Response.Write("document.getElementById('dpx').value=((nx+3)/2);"&vbcrlf)
    Response.Write("document.getElementById('dpy').value=((1023-ny-15)/2);"&vbcrlf)
    Response.Write("document.getElementById('did').value=o.id;"&vbcrlf)
	Response.Write("keyprBil.innerHTML = 'PosX: '+((nx+3)/2)+ ' PosY: ' + ((1023-ny-15)/2);"&vbcrlf)
	End If
		  %>
		  
          Drag.obj.root.onDrag(nx, ny);

          return false;
     },

     end : function()
     {
          document.onmousemove = null;
          document.onmouseup   = null;
          Drag.obj.root.onDragEnd(     parseInt(Drag.obj.root.style[Drag.obj.hmode ? "left" : "right"]), 
                                             parseInt(Drag.obj.root.style[Drag.obj.vmode ? "top" : "bottom"]));
			Drag.obj = null;
			npcsave();

     },

     fixE : function(e)
     {
          if (typeof e == 'undefined') e = window.event;
          if (typeof e.layerX == 'undefined') e.layerX = e.offsetX;
          if (typeof e.layerY == 'undefined') e.layerY = e.offsetY;
          return e;
     }
};

this.screenshotPreview = function(){	
	/* CONFIG */
		
		xOffset = 10;
		yOffset = 30;
		
		// these 2 variable determine popup's distance from the cursor
		// you might want to adjust to get the right result
		
	/* END CONFIG */
	$("img.screenshot").hover(function(e){
		this.t = this.title;
		this.title = "";	
		var c = (this.t != "") ? "" + this.t : "";
		$("body").append("<p id='screenshot'>"+ c +"</p>");								 
		$("#screenshot")
			.css("top",(e.pageY - xOffset) + "px")
			.css("left",(e.pageX + yOffset) + "px")
			.fadeIn("fast");						
    },
	function(){
		this.title = this.t;	
		$("#screenshot").remove();
    });	
	$("img.screenshot").mousemove(function(e){
		$("#screenshot")
			.css("top",(e.pageY - xOffset) + "px")
			.css("left",(e.pageX + yOffset) + "px");
	});			
};


// starting the script on page load
$(document).ready(function(){
	screenshotPreview();
});


function npcsave(){
$.ajax({
   type:'post',
   url: 'savenpc.asp',
   data:$('#npcbilgi').serialize()
   
	});
}

</script>
<style>
h1{
	font-size:180%;
	font-weight:normal;
	color:#555;
}
h2{
	clear:both;
	font-size:160%;
	font-weight:normal;
	color:#555;
	margin:0;
	padding:.5em 0;
}
a{
	text-decoration:none;
	color:#f30;	
}
p{
	clear:both;
	margin:0;
	padding:.5em 0;
}
pre{
	display:block;
	font:100% "Courier New", Courier, monospace;
	padding:10px;
	border:1px solid #bae2f0;
	background:#e3f4f9;	
	margin:.5em 0;
	overflow:auto;
	width:800px;
}

img{border:none;}
ul,li{
	margin:0;
	padding:0;
}
li{
	list-style:none;
	float:left;
	display:inline;
	margin-right:10px;
}
#screenshot{
	position:absolute;
	border:1px solid #ccc;
	background:#333;
	padding:5px;
	display:none;
	color:#fff;
	}
</style>
<body background="../imgs/WideScreenMaps/<%=zoneid%>.jpg" style="background-repeat:no-repeat">
<div style="width:1024px;height:1024px">
<img src="../imgs/Red_Arrow_Down.gif"  style="position:absolute;top:20px;display:none" id="okk">
<div id="keypressBilgi" style="position:fixed;left:520px;color:#fff"></div>
<form name="npcbilgi" id="npcbilgi">
<input type="hidden" id="dpx" name="dpx">
<input type="hidden" id="dpy" name="dpy">
<input type="hidden" id="did" name="did">
<input type="hidden" id="mapid" name="mapid" value="<%=zoneid%>"></form>
<%Do while not npcs.eof 
npcidx=npcs("id")
if zoneid="1" or zoneid="2" or zoneid="201" Then
pxx=(round(npcs("leftx")/2))-3
pyy=1023-(round(npcs("topz")/2))-15
elseif zoneid=21 Then
pxx=(round(npcs("leftx")*2))-3
pyy=1023-(round(npcs("topz")*2))-15
elseif zoneid="11" or zoneid="12" or zoneid="202" Then
pxx=(round(npcs("leftx")))-3
pyy=1023-(round(npcs("topz")))-15
End If
Response.Write "<div style=""position:absolute;left:"&pxx&";top:"&pyy&""" onDblClick=""lnpc('"&npcs("ssid")&"')""  id=""d"&npcidx&""" ><img src=""../imgs/Red_Arrow_Down.gif"" width=""7"" height=""15"" id=""i"&npcidx&""" class=""screenshot"" title="""&npcs("strname")&"("
If npcs("bygroup")=1 Then
Response.Write "<img src='../imgs/karuslogo.gif'>"
ElseIf npcs("bygroup")=2 Then
Response.Write "<img src='../imgs/elmologo.gif'>"
Else
Response.Write "<img src='../imgs/elmologo.gif'><img src='../imgs/karuslogo.gif'>"
End If
Response.Write ")<br>"&npcs("ssid")&"("&npcs("leftx")&","&npcs("topz")&")""></div>"&vbcrlf%>
<script>
     var d<%=npcidx%> = document.getElementById("d<%=npcidx%>");
     var i<%=npcidx%> = document.getElementById("i<%=npcidx%>");
  Drag.init(i<%=npcidx%>, <%="d"&npcidx%>);
</script>
<%npcs.MoveNext
Loop
npcs.Close
set npcs = Nothing %></div>
<script>
function lnpc(url) {
window.open('npcedit.asp?npcid='+url, 'Window2', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=500,height=500,top=0,left=0')
}
</script>
</body>