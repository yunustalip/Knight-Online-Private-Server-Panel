<!--#include file="../_inc/conn.asp"-->
<!--#include file="../Function.asp"--><link rel="stylesheet" href="../css/sIFR-screen.css" type="text/css" media="screen" />
<link rel="stylesheet" href="../css/sIFR-print.css" type="text/css" media="print" />
<script type="text/javascript" src="../js/sifr.js"></script>
<script type="text/javascript" src="../js/sifr-addons.js"></script>
<style>
h1{
	font-size: 50px;
	text-align:center;
	margin-top:25px;
	height:100px
}
</style>

<script language="JavaScript" type="text/JavaScript">
<%dim sitesettings
Set sitesettings = Conne.Execute("Select * From siteayar")

set FSO = Server.CreateObject("Scripting.FileSystemObject")
set Klasorler = FSO.GetFolder(Server.Mappath("../Font"))
Set Dosyalar = Klasorler.Files

For Each Dosya In Dosyalar  %>
//<![CDATA[
/* Replacement calls. Please see documentation for more information. */
if(typeof sIFR == "function"){
sIFR.replaceElement(named({sSelector:"h1#<%=replace(Dosya.name,".","")%>", sFlashSrc:"../font/<%=Dosya.name%>", sWmode: "transparent" , sColor:"#F9EED8",  nPaddingTop:55, nPaddingBottom:0, sFlashVars:"textalign=center&offsetTop=0"}));
};
//]]>
<%Next%>
</script>
<%For Each Dosya In Dosyalar 
Response.Write("<h1 id="""&replace(Dosya.name,".","")&""" >"&sitesettings("sitebaslik") &dosya.name&"</h1>"&vbcrlf)
Next%>