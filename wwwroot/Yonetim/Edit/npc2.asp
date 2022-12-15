<!--#include file="_inc/conn.asp"-->
<script src="js/jquery.js" type="text/javascript"></script>
<style>body{color:#000;font-size:12px}</style>
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

/*  */
</style><script type="text/javascript">
function keypressKont2(olay)
{
var keyprBil = document.getElementById('keypressBilgi');
olay = olay || event;
if (olay.clientX<512&&olay.clientY<512){
keyprBil.innerHTML = "PosX: "+ olay.clientX + " PosY: " + (511-olay.clientY);
document.getElementById('okk').style.left=olay.clientX-7+'px';
document.getElementById('okk').style.top=olay.clientY-20+'px';
document.getElementById('okk').style.display='inherit';

}
}

function keypressKont(olay)
{
var keyprBil = document.getElementById('keypressBilgi');
olay = olay || event;
if (olay.clientX<512&&olay.clientY<512){
keyprBil.innerHTML = "PosX: "+ olay.clientX + " PosY: " + (511-olay.clientY);

}
}
document.onmousedown= keypressKont2;
document.onmousemove= keypressKont;


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
</script>
<body background="imgs/Maps/201.jpg" style="background-repeat:no-repeat">
<img src="imgs/Red_Arrow_Down.gif"  style="position:absolute;top:20px;display:none" id="okk">
<div id="keypressBilgi" style="position:absolute;left:520px;"></div>
<%set npcs=Conne.Execute("select * from k_npcpos where zoneid=201")
do while not npcs.eof 
set npcname=Conne.Execute("select * from k_npc where ssid="&npcs("npcid"))
if not npcname.eof Then
pxx=(round(npcs("leftx")/4))-3
pyy=511-(round(npcs("topz")/4))-7
Response.Write "<div style=""position:absolute;left:"&pxx&";top:"&pyy&"""><img src=""imgs/Red_Arrow_Down.gif"" width=""7"" height=""15"" class=""screenshot"" title="""&npcname("strname")&"("
if npcname("bygroup")=1 Then
Response.Write "<img src=imgs/karuslogo.gif>"
elseif npcname("bygroup")=2 Then
Response.Write "<img src=imgs/elmologo.gif>"
else
Response.Write "<img src=imgs/elmologo.gif><img src=imgs/karuslogo.gif>"
End If
Response.Write ")<br>"&npcname("ssid")&"""></div>"
End If
npcs.movenext
loop%>