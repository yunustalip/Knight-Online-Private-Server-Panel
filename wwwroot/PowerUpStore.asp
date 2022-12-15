<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<% 
Dim MenuAyar,ksira,i,accountid,inventoryslot,ip,accid,itype,pus,sc,kiy,sil,tak,kal,adm,change,pusitems,s,pusitem,pusitem2,toplam_sonuc,git,ilkkayit
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then

accountid=secur(request.form("mgid"))
inventoryslot=secur(request.form("param"))
ip=Request.ServerVariables("REMOTE_ADDR")


If Session("login")="ok" And Not Session("username")=""  Then
Session("durum")="web"
ElseIf Session("login")<>"ok" And Session("username")="" And accountid<>"" and inventoryslot<>"" Then
Dim chrct,invslot
Set chrct=Conne.Execute("select * from currentuser where straccountid='"&accountid&"' and strclientip='"&ip&"'")

if chrct.eof Then
Session("durum")="web"
else
Session("durum")="game"
Session("accountid")=accountid
Session("charid")=chrct("strcharid")
invslot = split(inventoryslot,",")
Session("slot")=invslot(2)
End If

ElseIf Session("durum")="game" and Session("login")<>"ok" and Session("accountid")<>"" and Session("username")="" and accountid="" and inventoryslot="" Then
Set chrct=Conne.Execute("select * from currentuser where straccountid='"&Session("accountid")&"' and strclientip='"&ip&"'")

If chrct.Eof Then
Session("durum")="web"
Else
Session("charid")=chrct("strcharid")
End If

Else
Response.Redirect "hata.asp"
End If

If Session("durum")="game" and Session("accountid")<>"" Then
accid=Session("accountid")

ElseIf Session("login")="ok" and Session("username")<>"" Then
accid=Session("username")

Else

server.execute("hata.asp")
Response.End
End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head>
<title>Power Up Store</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-9">
<link href="images/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="js/form-submit.js"></script>
<script type="text/javascript" src="js/ajax.js"></script>
<script type="text/javascript" src="js/overlib.js"></script>
<script src="js/jquery.js" type="text/javascript"></script>
<script language="Javascript">

<!--

function buy_go(num) {
	right.location.href="buy.asp?itemid="+num;
}

function wish_go(num) {
	right.location.href="cart.asp?itemid="+num;
}

function detail_go(num) {
	right.location.href="detay.asp?itemid="+num;
}

var frm = document.myform;

function buy_go_premium(){
	frm.submit();
}

//-->
</script>
<script src="images/formFunctions.htm"></script>

</head><body link="#cccccc" vlink="#cccccc"   alink="#cccccc" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: url(images/bck.jpg); background-repeat: no-repeat;" onLoad="MM_preloadImages('images/butt/butt_buy_on.gif','images/butt/butt_detail_on.gif','images/butt/butt_cart_on.gif')" oncontextmenu="return false" ondragstart="return false" onselectstart="return false" >

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_CheckRecharge() 
{ //v3.0
  //alert("");

}

function recharge()
{
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="https://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="179" height="30" id="recharge" align="middle">');
	document.write('<param name="allowScriptAccess" value="sameDomain" />');
	document.write('<param name="movie" value="images/recharge.swf?clickTAG=powerupstore.asp?control=recharge" />');
	document.write('<param name="menu" value="false" />');
	document.write('<param name="quality" value="high" />');
	document.write('<param name="bgcolor" value="#000000" />');
	document.write('<embed src="images/recharge.swf?clickTAG=powerupstore.asp?control=recharge" menu="false" quality="high" bgcolor="#000000" width="179" height="30" name="recharge" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="https://www.macromedia.com/go/getflashplayer" />');
	document.write('</object>');
}

//-->
</script>

<table cellpadding="0" cellspacing="0" width="1024">
<tbody><tr>

    <td valign="middle" height="41">

    </td>
    <td align="right">
    <table border="0" cellpadding="2" cellspacing="0">
    <tbody><tr>
        <td>
        <script type="text/javascript">recharge();</script>

        </td>
        <td></td>
        <td width="8"></td>
   </tr>
   
    </tbody></table>
    </td>
</tr>
</tbody></table>

<div style="position: absolute; top: 115px; left: 30px;">
<table border="0" cellpadding="0" cellspacing="1">
<tbody><tr>
<% itype=secur(Request.Querystring("type"))
if itype="scroll" Then
sc="b"
elseif itype="kiyafet" Then
kiy="b"
elseif itype="silah" Then
sil="b"
elseif itype="taki" Then
tak="b"
elseif itype="kalkan" Then
kal="b"
elseif itype="admin" Then
adm="b"
else
pus="b"
End If
if itype="admin" and Session("yetki")="" Then
Response.Redirect("powerupstore.asp")
Response.End
End If%>

	<td><a href="powerupstore.asp"><img src="images/pus<%=pus%>.gif" id="pus" border="0" onMouseOver="MM_swapImage('pus','','images/pusb.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
	<td><a href="powerupstore.asp?type=scroll"><img src="images/scroll<%=sc%>.gif" id="scroll" border="0" onMouseOver="MM_swapImage('scroll','','images/scrollb.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
	<td><a href="powerupstore.asp?type=kiyafet"><img src="images/kiyafet<%=kiy%>.gif" id="kiyafet" border="0" onMouseOver="MM_swapImage('kiyafet','','images/kiyafetb.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
	<td><a href="powerupstore.asp?type=silah"><img src="images/silah<%=sil%>.gif" id="silah" border="0" onMouseOver="MM_swapImage('silah','','images/silahb.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
	<td><a href="powerupstore.asp?type=taki"><img src="images/taki<%=tak%>.gif" id="taki" border="0" onMouseOver="MM_swapImage('taki','','images/takib.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
	<td><a href="powerupstore.asp?type=kalkan"><img src="images/kalkan<%=kal%>.gif" id="kalkan" border="0" onMouseOver="MM_swapImage('kalkan','','images/kalkanb.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
<%if Session("yetki")="1" Then%>
	<td><a href="powerupstore.asp?type=admin"><img src="images/admin<%=adm%>.gif" id="admin" border="0" onMouseOver="MM_swapImage('admin','','images/adminb.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
<%End If%>
</tr>
</tbody></table>
</div>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>


<body onLoad="MM_preloadImages('images/butt/butt_detail_on.gif','images/butt/butt_cart_on.gif','images/butt/butt_buy_on.gif')">
<div style="position: absolute; top: 170px; left: 43px;">
      <table border="0" cellpadding="0" cellspacing="0">

        <tbody><tr>
          <td align="center" height="555" valign="top"><table border="0" cellpadding="0" cellspacing="0">
            <tbody><tr>
                <td align="left" height="20">
                

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            </td>
              </tr>

              <tr>
  <td>

 <table border="0" cellpadding="0" cellspacing="0" width="635">

<tbody><tr>
<%if secur(Request.Querystring("control"))="recharge" Then
server.execute("recharge.asp")
Response.End

elseif secur(Request.Querystring("control"))="recharge2" Then
server.execute("recharge.asp")
Response.End
End If

if itype="" Then

change=secur(Request.Querystring("change"))%>
<style type="text/css">
<!--
.style3 {color: #999999; font-weight: bold; }
.style4 {
	font-family: Arial;
	font-size: 12px;
}
.style5 {
	color: #CCCCCC;
	font-weight: bold;
}
-->
</style>
<body onLoad="MM_preloadImages('images/butt/butt_detail_on.gif','images/butt/butt_buy_on.gif')">
<% if change=""  Then%>
<img src="images/pus-premium-sale.gif" width="631" height="293" /> <br /><br />
<table cellpadding="0" cellspacing="0">
<tr>
<% set pusitems=Conne.Execute("select * from pus_itemleri where type='anasayfa'")
s=1 
do while not pusitems.eof%>
<td>
<table background="images/n_table_01.gif" border="0" cellpadding="0" cellspacing="0" height="155" width="213">
 <tbody>
 <tr>
 <td valign="top">
 	<table border="0" style="display:" cellpadding="0" cellspacing="0" width="200">
 	<tbody>
 	<tr align="center" valign="middle">
 	<td class="12_yellow" width="20">&nbsp;&nbsp;</td>
 	<td colspan="2" style="font-family: 'verdana',Times,serif; color: rgb(0, 0, 0); font-size: 10px; padding-top: 10px;" height="30"><strong>* <%=pusitems("itemismi")%></strong></td>
 	</tr>
    <tr>
    <td align="right" width="78">&nbsp;</td>
    <td align="right" width="78"><table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tbody><tr>
    <td align="right" height="80" valign="bottom">
	<table border="0" cellpadding="0" cellspacing="0" height="72" width="72">
    <tbody><tr>
    <td align="center" background="images/table1_01.gif"><a href="powerupstore.asp?change=<%=pusitems("detay")%>"><img src="item/<%=pusitems("resim")%>" width="64" height="64" border="0" /></a></td>
     </tr></tbody></table>
	 </td></tr>
     <tr>
     <td align="center" height="32" valign="bottom">&nbsp;</td>
     </tr>
     </tbody></table></td>
     <td align="center" valign="bottom" width="102"><table border="0" cellpadding="0" cellspacing="0" height="108" width="106">

     <tbody><tr>
     <td align="right"><table border="0" cellpadding="0" cellspacing="0" width="100">
     <tbody><tr>
     <td style="font-family: 'verdana',Times,serif; color: rgb(0, 0, 0); font-size: 10px;" width="10">-</td>
     <td style="font-family: verdana,sans-serif; color: rgb(0, 0, 0); font-size: 11px;" height="16">Ücret : <%=pusitems("ucret")%></td>
        </tr>
        </tbody></table></td>
        </tr>
        <tr>
        <td align="center" height="30">
       <a href="powerupstore.asp?change=<%=pusitems("detay")%>" onMouseOver="MM_swapImage('butt_buy<%=s%>','','images/butt/butt_buy_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt//butt_buy.gif" name="butt_buy<%=s%>" width="89" height="30" border="0" id="butt_buy<%=s%>" /></a>
	    <a href="javascript:buy_go('<%=s%>');" onMouseOver="MM_swapImage('butt_buy1','','images/butt/butt_buy_on.gif',1)" onMouseOut="MM_swapImgRestore()"></a></td>
                                      </tr>
                                      <tr>
                                        <td align="center" height="30"></td>
                                    </tr>
                                  </tbody></table></td>
                                </tr>
                              </tbody></table>
		
                            </td>

                          </tr>
                        </tbody></table>                        </td>
<%
s=s+1
pusitems.movenext
loop %>
</tr></table>
<%elseif change="nation" or change="nation2"  Then
Dim ucht,accchr,sql
	Set ucht = Server.CreateObject("ADODB.Recordset")
	sql = "Select * From tb_user where strAccountID='"&accid&"'"
	ucht.open sql,conne,1,3
	
	if ucht.eof Then
	Session.abandon
	Response.Redirect("default.asp")
	Response.End
	End If
	
	Set accchr = Server.CreateObject("ADODB.Recordset")
	sql = "Select * From ACCOUNT_CHAR where strAccountID='"&accid&"'"
	accchr.open sql,conne,1,3

	if not accchr.eof Then
	Dim onlinekontrol
	set onlinekontrol=Conne.Execute("select * from CURRENTUSER where straccountid='"&accid&"'")
	if not onlinekontrol.eof Then
	Response.Write "<center>Lütfen Oyundan Çýkýþ Yapýp Ýþleminizi WebSite Üzerinden Devam Ettiriniz.</center>"
	Response.End
	End If
	Dim mycharst1,mycharst2,mycharst3
	Set mycharst1 = Server.CreateObject("ADODB.Recordset")
	sql = "Select * From USERDATA where struserID='"&accchr("strCharID1")&"'"
	mycharst1.open sql,conne,1,3

	Set mycharst2 = Server.CreateObject("ADODB.Recordset")
	sql = "Select * From USERDATA where struserID='"&accchr("strCharID2")&"'"
	mycharst2.open sql,conne,1,3

	Set mycharst3 = Server.CreateObject("ADODB.Recordset")
	sql = "Select * From USERDATA where struserID='"&accchr("strCharID3")&"'"
	mycharst3.open sql,conne,1,3

	if change="nation" Then %>
	<br><center><font face="times new roman" class="style3">NATION CHANGE</font></center>
	<ul><li><font class="style4">Job veya race lerde hatalar oluþmasý durumunda <a href="default.asp?cat=submitticket">Ticket atýnýz</a><font></li>
	</ul>
	<br/><br/><center>
	<font class="style4"><b>Irk Deðiþtirmekte Son Kararýnýz Mý ? </b><br />
	
	<a href="powerupstore.asp?change=nation2">Evet</a> | <a href="powerupstore.asp">Hayýr</a></font>
	 <% elseif change="nation2" Then 
	 Dim cashkontrol,nationcash,q
	set cashkontrol=Conne.Execute("select cashpoint from tb_user where straccountid='"&accid&"'")
	set nationcash=Conne.Execute("select ucret from pus_itemleri where itemkodu='800090000'")
	
	if not cashkontrol("cashpoint")>=nationcash("ucret") Then
	Response.Write("<center><br>Yetersiz Cash Point !<br><br>Mevcut Cash Point: "&cashkontrol("cashpoint")&"<br>Gerekli Cashpoint: "&nationcash("ucret")&"</center>")
	Response.End
	End If
	if accchr("bNation")="1" Then
	q="2"
	elseif accchr("bNation")="2" Then
	q="1"
	else
	Response.End
	End If
	
	if q="2" Then 
	
	if not mycharst1.eof Then
	select case mycharst1("Race")
	case "1"
	mycharst1("Race")="11"
	case "2"
	mycharst1("Race")="12"
	case "3"
	mycharst1("Race")="13"
	case "4"
	mycharst1("Race")="13"
	end select
	mycharst1("Class")=mycharst1("Class")+100
	mycharst1("Nation")=q
	mycharst1.update
	
	if mycharst1("knights") <> "0" AND mycharst1("fame") = "2" Then
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1 = NULL WHERE ViceChief_1 = '"&mycharst1("struserid")&"' ")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2 = NULL WHERE ViceChief_2 = '"&mycharst1("struserid")&"' ")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3 = NULL WHERE ViceChief_3 = '"&mycharst1("struserid")&"' ")
	elseif mycharst1("knights") <> "0" AND mycharst1("fame") = "1" Then
	Conne.Execute("DELETE FROM KNIGHTS WHERE IDNum ='"&mycharst1("knights")&"' ")
	Conne.Execute("DELETE FROM KNIGHTS_USER WHERE sIDNum ='"&mycharst1("knights")&"' ")
	Conne.Execute("UPDATE USERDATA SET Knights = 0, Fame = 0 WHERE Knights ='"&mycharst1("knights")&"' ")
	End If
	
	Conne.Execute("Delete from KNIGHTS_USER where struserid='"&mycharst1("struserid")&"' ")
	Conne.Execute("update KNIGHTS set Members=Members-1 where idnum='"&mycharst1("knights")&"' ")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName = NULL WHERE strKingName='"&mycharst1("struserid")&"'")
	Conne.Execute("UPDATE USERDATA set Knights=0, Fame=0, Rank=0, Title=0, Zone=21,px=31200, pz=40200,py=0 where struserid='"&mycharst1("struserid")&"' ")
	End If
	
	if not mycharst2.eof Then
	select case mycharst2("Race")
	case "1"
	mycharst2("Race")="11"
	case "2"
	mycharst2("Race")="12"
	case "3"
	mycharst2("Race")="13"
	case "4"
	mycharst2("Race")="13"
	end select
	mycharst2("Class")=mycharst2("Class")+100
	mycharst2("Nation")=q
	mycharst2.update
	
	if mycharst2("knights") <> "0" AND mycharst2("fame") = "2" Then
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1 = NULL WHERE ViceChief_1 = '"&mycharst2("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2 = NULL WHERE ViceChief_2 = '"&mycharst2("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3 = NULL WHERE ViceChief_3 = '"&mycharst2("struserid")&"'")
	elseif mycharst2("knights") <> "0" AND mycharst2("fame") = "1" Then
	Conne.Execute("DELETE FROM KNIGHTS WHERE IDNum ='"&mycharst2("knights")&"' ")
	Conne.Execute("DELETE FROM KNIGHTS_USER WHERE sIDNum ='"&mycharst2("knights")&"' ")
	Conne.Execute("UPDATE USERDATA SET Knights = 0, Fame = 0 WHERE Knights ='"&mycharst2("knights")&"' ")
	End If
	Conne.Execute("Delete from KNIGHTS_USER where struserid='"&mycharst2("struserid")&"' ")
	Conne.Execute("update KNIGHTS set Members=Members-1 where idnum='"&mycharst2("knights")&"' ")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName = NULL WHERE strKingName='"&mycharst2("struserid")&"'")
	Conne.Execute("UPDATE USERDATA set Knights=0, Fame=0, Rank=0, Title=0, Zone=21,px=31200, pz=40200,py=0 where struserid='"&mycharst2("struserid")&"' ")
	End If
	
	if not mycharst3.eof Then
	select case mycharst3("Race")
	case "1"
	mycharst3("Race")="11"
	case "2"
	mycharst3("Race")="12"
	case "3"
	mycharst3("Race")="13"
	case "4"
	mycharst3("Race")="13"
	end select
	mycharst3("Class")=mycharst3("Class")+100
	mycharst3("Nation")=q
	mycharst3.update
	
	if mycharst3("knights") <> "0" AND mycharst3("fame") = "2" Then
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1 = NULL WHERE ViceChief_1 = '"&mycharst3("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2 = NULL WHERE ViceChief_2 = '"&mycharst3("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3 = NULL WHERE ViceChief_3 = '"&mycharst3("struserid")&"'")
	elseif mycharst3("knights") <> "0" AND mycharst3("fame") = "1" Then
	Conne.Execute("DELETE FROM KNIGHTS WHERE IDNum ='"&mycharst3("knights")&"' ")
	Conne.Execute("DELETE FROM KNIGHTS_USER WHERE sIDNum ='"&mycharst3("knights")&"' ")
	Conne.Execute("UPDATE USERDATA SET Knights = 0, Fame = 0 WHERE Knights ='"&mycharst3("knights")&"' ")
	End If
	
	Conne.Execute("Delete from KNIGHTS_USER where struserid='"&mycharst3("struserid")&"' ")
	Conne.Execute("update KNIGHTS set Members=Members-1 where idnum='"&mycharst3("knights")&"' ")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName = NULL WHERE strKingName='"&mycharst3("struserid")&"'")
	Conne.Execute("UPDATE USERDATA set Knights=0, Fame=0, Rank=0, Title=0, Zone=21,px=31200 pz=40200,py=0 where struserid='"&mycharst3("struserid")&"' ")
	End If
	
	elseif q="1" Then 
	
	if not mycharst1.eof Then
	select case mycharst1("Race")
	case "11"
	mycharst1("Race")="1"
	case "12"
	if mycharst1("Class")="201" or mycharst1("Class")="205" or mycharst1("Class")="206" Then
	mycharst1("Race")="1"
	elseif mycharst1("Class")="202" or mycharst1("Class")="207" or mycharst1("Class")="208" Then
	mycharst1("Race")="2"
	elseif mycharst1("Class")="203" or mycharst1("Class")="209" or mycharst1("Class")="210" Then
	mycharst1("Race")="3"
	elseif mycharst1("Class")="204" or mycharst1("Class")="211" or mycharst1("Class")="212" Then
	mycharst1("Race")="2"
	End If
	case "13"
	if mycharst1("Class")="201" or mycharst1("Class")="205" or mycharst1("Class")="206" Then
	mycharst1("Race")="1"
	elseif mycharst1("Class")="202" or mycharst1("Class")="207" or mycharst1("Class")="208" Then
	mycharst1("Race")="2"
	elseif mycharst1("Class")="203" or mycharst1("Class")="209" or mycharst1("Class")="210" Then
	mycharst1("Race")="3"
	elseif mycharst1("Class")="204" or mycharst1("Class")="211" or mycharst1("Class")="212" Then
	mycharst1("Race")="4"
	End If	
	end select
	mycharst1("Class")=mycharst1("Class")-100
	mycharst1("Nation")=q
	mycharst1.update
	
	if mycharst1("knights") <> "0" AND mycharst1("fame") = "2" Then
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1 = NULL WHERE ViceChief_1 = '"&mycharst1("struserid")&"' ")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2 = NULL WHERE ViceChief_2 = '"&mycharst1("struserid")&"' ")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3 = NULL WHERE ViceChief_3 = '"&mycharst1("struserid")&"' ")
	elseif mycharst1("knights") <> "0" AND mycharst1("fame") = "1" Then
	Conne.Execute("DELETE FROM KNIGHTS WHERE IDNum ='"&mycharst1("knights")&"' ")
	Conne.Execute("DELETE FROM KNIGHTS_USER WHERE sIDNum ='"&mycharst1("knights")&"' ")
	Conne.Execute("UPDATE USERDATA SET Knights = 0, Fame = 0 WHERE Knights ='"&mycharst1("knights")&"' ")
	End If
	
	Conne.Execute("Delete from KNIGHTS_USER where struserid='"&mycharst1("struserid")&"' ")
	Conne.Execute("update KNIGHTS set Members=Members-1 where idnum='"&mycharst1("knights")&"' ")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName = NULL WHERE strKingName='"&mycharst1("struserid")&"'")
	Conne.Execute("UPDATE USERDATA set Knights=0, Fame=0, Rank=0, Title=0, Zone=21 where struserid='"&mycharst1("struserid")&"' ")
	End If
	
	if not mycharst2.eof Then
	select case mycharst2("Race")
	case "11"
	mycharst2("Race")="1"
	case "12"
	if mycharst2("Class")="201" or mycharst2("Class")="205" or mycharst2("Class")="206" Then
	mycharst2("Race")="1"
	elseif mycharst2("Class")="202" or mycharst2("Class")="207" or mycharst2("Class")="208" Then
	mycharst2("Race")="2"
	elseif mycharst2("Class")="203" or mycharst2("Class")="209" or mycharst2("Class")="210" Then
	mycharst2("Race")="3"
	elseif mycharst2("Class")="204" or mycharst2("Class")="211" or mycharst2("Class")="212" Then
	mycharst2("Race")="2"
	End If	
	case "13"
	if mycharst2("Class")="201" or mycharst2("Class")="205" or mycharst2("Class")="206" Then
	mycharst1("Race")="1"
	elseif mycharst2("Class")="202" or mycharst2("Class")="207" or mycharst2("Class")="208" Then
	mycharst1("Race")="2"
	elseif mycharst2("Class")="203" or mycharst2("Class")="209" or mycharst2("Class")="210" Then
	mycharst1("Race")="3"
	elseif mycharst2("Class")="204" or mycharst2("Class")="211" or mycharst2("Class")="212" Then
	mycharst2("Race")="2"
	End If
	end select
	mycharst2("Class")=mycharst2("Class")-100
	mycharst2("Nation")=q
	mycharst2.update
	
	if mycharst2("knights") <> "0" AND mycharst2("fame") = "2" Then
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1 = NULL WHERE ViceChief_1 = '"&mycharst2("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2 = NULL WHERE ViceChief_2 = '"&mycharst2("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3 = NULL WHERE ViceChief_3 = '"&mycharst2("struserid")&"'")
	elseif mycharst2("knights") <> "0" AND mycharst2("fame") = "1" Then
	Conne.Execute("DELETE FROM KNIGHTS WHERE IDNum ='"&mycharst2("knights")&"' ")
	Conne.Execute("DELETE FROM KNIGHTS_USER WHERE sIDNum ='"&mycharst2("knights")&"' ")
	Conne.Execute("UPDATE USERDATA SET Knights = 0, Fame = 0 WHERE Knights ='"&mycharst2("knights")&"' ")
	End If
	Conne.Execute("Delete from KNIGHTS_USER where struserid='"&mycharst2("struserid")&"' ")
	Conne.Execute("update KNIGHTS set Members=Members-1 where idnum='"&mycharst2("knights")&"' ")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName = NULL WHERE strKingName='"&mycharst2("struserid")&"'")
	Conne.Execute("UPDATE USERDATA set Knights=0, Fame=0, Rank=0, Title=0, Zone=21 where struserid='"&mycharst2("struserid")&"' ")
	End If
	
	if not mycharst3.eof Then
	select case mycharst3("Race")
	case "11"
	mycharst3("Race")="1"
	case "12"
	if mycharst3("Class")="201" or mycharst3("Class")="205" or mycharst3("Class")="206" Then
	mycharst3("Race")="1"
	elseif mycharst3("Class")="202" or mycharst3("Class")="207" or mycharst3("Class")="208" Then
	mycharst3("Race")="2"
	elseif mycharst3("Class")="203" or mycharst3("Class")="209" or mycharst3("Class")="210" Then
	mycharst3("Race")="3"
	elseif mycharst3("Class")="204" or mycharst3("Class")="211" or mycharst3("Class")="212" Then
	mycharst3("Race")="2"
	End If
	case "13"
	if mycharst3("Class")="201" or mycharst3("Class")="205" or mycharst3("Class")="206" Then
	mycharst3("Race")="1"
	elseif mycharst3("Class")="202" or mycharst3("Class")="207" or mycharst3("Class")="208" Then
	mycharst3("Race")="2"
	elseif mycharst3("Class")="203" or mycharst3("Class")="209" or mycharst3("Class")="210" Then
	mycharst3("Race")="3"
	elseif mycharst3("Class")="204" or mycharst3("Class")="211" or mycharst3("Class")="212" Then
	mycharst3("Race")="2"
	End If
	end select
	mycharst3("Class")=mycharst3("Class")-100
	mycharst3("Nation")=q
	mycharst3.update
	if mycharst3("knights") <> "0" AND mycharst3("fame") = "2" Then
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1 = NULL WHERE ViceChief_1 = '"&mycharst3("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2 = NULL WHERE ViceChief_2 = '"&mycharst3("struserid")&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3 = NULL WHERE ViceChief_3 = '"&mycharst3("struserid")&"'")
	elseif mycharst3("knights") <> "0" AND mycharst3("fame") = "1" Then
	Conne.Execute("DELETE FROM KNIGHTS WHERE IDNum ='"&mycharst3("knights")&"' ")
	Conne.Execute("DELETE FROM KNIGHTS_USER WHERE sIDNum ='"&mycharst3("knights")&"' ")
	Conne.Execute("UPDATE USERDATA SET Knights = 0, Fame = 0 WHERE Knights ='"&mycharst3("knights")&"' ")
	End If
	
	Conne.Execute("Delete from KNIGHTS_USER where struserid='"&mycharst3("struserid")&"' ")
	Conne.Execute("update KNIGHTS set Members=Members-1 where idnum='"&mycharst3("knights")&"' ")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName = NULL WHERE strKingName='"&mycharst3("struserid")&"'")
	Conne.Execute("UPDATE USERDATA set Knights=0, Fame=0, Rank=0, Title=0, Zone=21 where struserid='"&mycharst3("struserid")&"' ")
	End If
	
	
	
	End If
	
	Dim qnation,qntn,ips
	q=accchr("bNation")
	accchr.update
	if q="1" Then
	qnation="Karus"
	elseif q="2" Then
	qnation="ElMorad"
	End If

	if accchr("bNation")="1" Then
	qntn="Karus"
	elseif accchr("bNation")="2" Then
	qntn="ElMorad"
	End If

	Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&accid&" nickli karakter "&qntn&" dan "&qnation&" a Irk transferi yaptý.','"&now&"')")

	%>
	<meta http-equiv="refresh" content="2;url=powerupstore.asp" />
	Irk <% if q="1" Then
	Response.Write"<b>Karus</b> olarak deðiþtirildi."
	elseif q="2" Then
	Response.Write"<b>Human</b> olarak deðiþtirildi."
	End If
	End If %>
	</center>
	
	<%
	else 
	Response.Write("Irk Deðiþtirebilmeniz için en az 1 Karakteriniz Olmasý Gerekiyor")
	End If 
	
	ucht.close
	set ucht=nothing
	accchr.close
	set accchr=nothing
	mycharst1.close
	set mycharst1=nothing
	mycharst2.close
	set mycharst2=nothing
	mycharst3.close
	set mycharst3=nothing
	
	
elseif change="name" Then
 if Session("login")="ok" Then
 Dim char
  set char=Conne.Execute("select * from account_char where straccountid='"&accid&"'")
  if not char.eof Then
	%><center>
	<form action="powerupstore.asp?change=namechangeok" method="post">
    <table width="300" border="0" align="center">
    <tr>
    <td colspan="2" align="center" background="imgs/menubg.gif"><font class="style5" size="3">Name Change</font></td>
    </tr>
  <tr>
    <td><font class="style4">Karekteri Seçiniz :</font></td>
    <td ><select name="charid" class="styleform">
    <% if len(trim(char("strcharid1")))>0 Then
	Response.Write "<option value='"&char("strcharid1")&"'>"&char("strcharid1")&"</option>"
	End If
	if len(trim(char("strcharid2")))>0 Then
	Response.Write "<option value='"&char("strcharid2")&"'>"&char("strcharid2")&"</option>"
	End If
	if len(trim(char("strcharid3")))>0 Then
	Response.Write "<option value='"&char("strcharid3")&"'>"&char("strcharid3")&"</option>"
	End If %>
    </select></td>
  </tr>
  <tr>
    <td><font class="style4">Yeni Ismi Giriniz. :</font></td>
    <td><input type="text" name="newid"></td>
  </tr>
  <tr><td colspan="2" align="center"><input type="submit" value="Deðiþtir" style="font-family:Tahoma; font-size: 12px;" onClick="this.value='Deðiþtiriliyor...';this.form.submit()"></td>
  </tr>
</table>
    </form></center>
	<%else
	Response.Write("Name change için en az 1 karakteriniz bulunmalýdýr.")
	End If
	else
	Response.Redirect("default.asp")
	End If
	elseif change="namechangeok" Then
	if Session("login")="ok" Then
	set cashkontrol=Conne.Execute("select cashpoint from tb_user where straccountid='"&accid&"'")
	Dim namecash,charid,newid,x,onkontrol,charkontrol,newidkontrol
	set namecash=Conne.Execute("select ucret from pus_itemleri where itemkodu='999999998'")
	
	if not cashkontrol("cashpoint")>=namecash("ucret") Then
	Response.Write("<center><br>Yetersiz Cash Point !<br><br>Mevcut cashpoint: "&cashkontrol("cashpoint")&"<br>Gerekli Cashpoint: "&namecash("ucret")&"</center>")
	Response.End
	End If
	


	charid=trim(secur(request.form("charid")))
	newid=trim(secur(request.form("newid")))

for x=1 to len(newid)
if mid(newid,x,1)="A" or mid(newid,x,1)="B" or mid(newid,x,1)="C" or mid(newid,x,1)="D" or mid(newid,x,1)="E" or mid(newid,x,1)="F" or mid(newid,x,1)="G" or mid(newid,x,1)="H" or mid(newid,x,1)="I" or mid(newid,x,1)="J" or mid(newid,x,1)="K" or mid(newid,x,1)="L" or mid(newid,x,1)="M" or mid(newid,x,1)="N" or mid(newid,x,1)="O" or mid(newid,x,1)="P" or mid(newid,x,1)="R" or mid(newid,x,1)="S" or mid(newid,x,1)="T" or mid(newid,x,1)="U" or mid(newid,x,1)="V" or mid(newid,x,1)="Y" or mid(newid,x,1)="Z" or mid(newid,x,1)="X" or mid(newid,x,1)="Q" or mid(newid,x,1)="W" or mid(newid,x,1)="a" or mid(newid,x,1)="b" or mid(newid,x,1)="c" or mid(newid,x,1)="d" or mid(newid,x,1)="e" or mid(newid,x,1)="f" or mid(newid,x,1)="g" or mid(newid,x,1)="h" or mid(newid,x,1)="i" or mid(newid,x,1)="j" or mid(newid,x,1)="k" or mid(newid,x,1)="l" or mid(newid,x,1)="m" or mid(newid,x,1)="n" or mid(newid,x,1)="o" or mid(newid,x,1)="p" or mid(newid,x,1)="r" or mid(newid,x,1)="s" or mid(newid,x,1)="t" or mid(newid,x,1)="u" or mid(newid,x,1)="v" or mid(newid,x,1)="y" or mid(newid,x,1)="z" or mid(newid,x,1)="x" or mid(newid,x,1)="q" or mid(newid,x,1)="w" or mid(newid,x,1)="0" or mid(newid,x,1)="1" or mid(newid,x,1)="2" or mid(newid,x,1)="3" or mid(newid,x,1)="4" or mid(newid,x,1)="5" or mid(newid,x,1)="6" or mid(newid,x,1)="7" or mid(newid,x,1)="8" or mid(newid,x,1)="9" Then

else
Response.Write "<center><font color=red><b>Lütfen özel karakterler Kullanmayýnýz!</b></font>"
Response.End 
End If
next

	if not len(charid)>0 Then
	Response.Write("<center><br><b>Lütfen Geçerli Bir Nick Giriniz..</b></center>")
	Response.End
	End If


	set onkontrol=Conne.Execute("select * from currentuser where strcharid='"&charid&"'")
	if not onkontrol.eof Then
	Response.Write "Lütfen oyundan çýktýktan sonra iþleminizi gerçekleþtiriniz."
	else
	set charkontrol=Conne.Execute("select * from account_char where straccountid='"&accid&"' and strcharid1='"&charid&"' or strcharid2='"&charid&"' or strcharid3='"&charid&"'")
	if charkontrol.eof Then
	Response.Write ("Karekter Bulunamadý !!!")&Response.End
	End If
	set newidkontrol=Conne.Execute("select * from userdata where struserid='"&newid&"'")
	if not newidkontrol.eof Then
	Response.Write "<center>Veritabanýnda &nbsp;"&newid&"&nbsp; adýnda karekter bulunmaktadýr. Lütfen Baþka isim deneyin<br><a href='javascript:history.back(-1)'>Geri Dön</center>" 
	else
	Conne.Execute("update account_char set strcharid1='"&newid&"' where strcharid1='"&charid&"'")
	Conne.Execute("update account_char set strcharid2='"&newid&"' where strcharid2='"&charid&"'")
	Conne.Execute("update account_char set strcharid3='"&newid&"' where strcharid3='"&charid&"'")
	Conne.Execute("update userdata set struserid='"&newid&"' where struserid='"&charid&"'")	
	Conne.Execute("UPDATE KNIGHTS_USER SET strUserId='"&newid&"' where strUserId='"&charid&"'")
	Conne.Execute("UPDATE KNIGHTS SET Chief='"&newid&"' where Chief='"&charid&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_1='"&newid&"' where ViceChief_1='"&charid&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_2='"&newid&"' where ViceChief_2='"&charid&"'")
	Conne.Execute("UPDATE KNIGHTS SET ViceChief_3='"&newid&"' where ViceChief_3='"&charid&"'")
	Conne.Execute("UPDATE KING_SYSTEM SET strKingName='"&newid&"' where strKingName='"&charid&"'")
	Conne.Execute("UPDATE KING_ELECTION_LIST SET strName='"&newid&"' where strName='"&charid&"'")
	Conne.Execute("UPDATE CURRENTUSER SET strCharID = '"&newid&"' where strCharID = '"&charid&"'")
	Conne.Execute("UPDATE USERDATA_SKILLSHORTCUT SET strCharID = '"&newid&"' WHERE strCharID = '"&charid&"' ")
	Conne.Execute("UPDATE USER_SAVED_MAGIC SET strCharID = '"&newid&"' WHERE strCharID = '"&charid&"' ")
	Conne.Execute("UPDATE FRIEND_LIST SET strUserID = '"&newid&"' where strUserID = '"&charid&"'")
	Conne.Execute("update tb_user set cashpoint=cashpoint-"&namecash("ucret")&" where straccountid='"&accid&"'")
	Conne.Execute("update pus_itemleri set alindi=alindi+1 where itemkodu='800032000'")
	Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter nickini "&newid&" olarak deðiþtirdi.','"&now&"')")
	Response.Write charid&"&nbsp;olan karakterinizin adý&nbsp;<b>"&newid&"</b>&nbsp; olarak baþarýyla deðiþtirilmiþtir."%>
	<meta http-equiv="refresh" content="2;url=powerupstore.asp" />
	<%End If
	End If
	else
	Response.Redirect("powerupstore.asp")
	End If
elseif change="job" Then
set chars=Conne.Execute("select * from account_char where straccountid='"&accid&"'")
if not chars.eof Then
char1=trim(chars("strcharid1"))
char2=trim(chars("strcharid2"))
char3=trim(chars("strcharid3"))%>
<form action="powerupstore.asp?change=job2" method="post">
  <table width="550" height="195" align="center">
  <tr><td colspan="4" align="center"><span class="style3">JOB CHANGE</span></td>
  </tr>
  <tr><td colspan="4" align="left">
    <li class="style4">Ýþlemlerinizi gerçekleþtirmeden önce üstünüzdeki itemleri bankanýza koyun.(Üstünüzdeki bütün itemleriniz silinecektir) </li>
    <li class="style4">Ardýndan Oyundan Çýkýn ! </li></td>
  </tr>
<tr>
<td width="95" rowspan="2" align="right" valign="middle">&nbsp;</td>
<td width="155" valign="middle" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px">Karakterinizi seçiniz</td>
<td width="200"  valign="middle" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px">Yeni Class Seçiniz</td>
<td>&nbsp;</td>
</tr>
<tr>
<td>
<select name="character" style=" font:Arial; font-size:12px">
<%if len(char1)>0 Then
Response.Write "<option value=""1"">"&char1&"</option>"&vbcrlf
End If
if len(char2)>0 Then
Response.Write "<option value=""2"">"&char2&"</option>"&vbcrlf
End If
if len(char3)>0 Then
Response.Write "<option value=""3"">"&char3&"</option>"&vbcrlf
End If%>
</select>
</td>
<td><select name="class" style=" font:Arial; font-size:12px">
    <option value="warrior">Warrior</option>
    <option value="rogue">Rogue </option>
    <option value="mage">Mage</option>
    <option value="priest">Priest</option>
  </select></td>
<td> <input type="submit" value="Deðiþtir&gt;&gt;" style="font:Arial; font-size:12px; cursor:pointer">
</td>
</tr>
  </table>
</form>
<%
End If
elseif change="job2" Then

set onlinec=Conne.Execute("select * from currentuser where straccountid='"&accid&"' and strcharid='"&secur(request.Form("character"))&"'")
if not onlinec.eof Then
Response.Write "<center>Lütfen Karakterinizi Oyundan Çýkardýktan Sonra Ýþleminize Devam Ediniz.</center>"
Response.End
End If

set cash=Conne.Execute("select cashpoint,straccountid from tb_user where straccountid='"&accid&"'")
if cash.eof Then
Response.End
End If
set jobcash=Conne.Execute("select itemkodu,ucret from pus_itemleri where itemkodu=999999999")
if cash("cashpoint")<jobcash("ucret") Then
Response.Write("<center><br>Yetersiz Cash Point !<br><br>Mevcut cashpoint: "&cash("cashpoint")&"<br>Gerekli Cashpoint: "&jobcash("ucret")&"</center>")
Response.End
End If
set chars=Conne.Execute("select * from account_char where straccountid='"&accid&"'")
if not chars.eof Then
char1=trim(chars("strcharid1"))
char2=trim(chars("strcharid2"))
char3=trim(chars("strcharid3"))

charid=secur(request.Form("character"))
charclass=secur(request.Form("class"))
If charid="1" or charid="2" or charid="3" Then
dim karakterb,bul
bul="strCharId"&cint(charid)

set karakterb=Conne.Execute("Select "&bul&" From Account_Char Where StrAccountID='"&accid&"'")
charid=karakterb(0)
If chars("bnation")="1" Then

Set charc=Conne.Execute("select * from userdata where struserid='"&charid&"'")
if not charc.eof Then

If charclass="warrior" Then
select case charc("class")
case "101"
newclass="101"
newrace="1"
case "105"
newclass="105"
newrace="1"
case "106"
newclass="106"
newrace="1"

case "102"
newclass="101"
newrace="1"
case "107"
newclass="105"
newrace="1"
case "108"
newclass="106"
newrace="1"

case "103"
newclass="101"
newrace="1"
case "109"
newclass="105"
newrace="1"
case "110"
newclass="106"
newrace="1"

case "104"
newclass="101"
newrace="1"
case "111"
newclass="105"
newrace="1"
case "112"
newclass="106"
newrace="1"
end select
Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
ElseIf charclass="rogue" Then
select case charc("class")
case "101"
newclass="102"
newrace="2"
case "105"
newclass="107"
newrace="2"
case "106"
newclass="108"
newrace="2"

case "102"
newclass="102"
newrace="2"
case "107"
newclass="107"
newrace="2"
case "108"
newclass="108"
newrace="2"

case "103"
newclass="102"
newrace="2"
case "109"
newclass="107"
newrace="2"
case "110"
newclass="108"
newrace="2"

case "104"
newclass="102"
newrace="2"
case "111"
newclass="107"
newrace="2"
case "112"
newclass="108"
newrace="2"
end select
Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
ElseIf charclass="mage" Then
select case charc("class")
case "101"
newclass="103"
newrace="3"
case "105"
newclass="109"
newrace="3"
case "106"
newclass="110"
newrace="3"

case "102"
newclass="103"
newrace="3"
case "107"
newclass="109"
newrace="3"
case "108"
newclass="110"
newrace="3"

case "103"
newclass="103"
newrace="3"
case "109"
newclass="109"
newrace="3"
case "110"
newclass="110"
newrace="3"

case "104"
newclass="103"
newrace="3"
case "111"
newclass="109"
newrace="3"
case "112"
newclass="110"
newrace="3"
end select
Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
ElseIf charclass="priest" Then
select case charc("class")
case "101"
newclass="104"
newrace="4"
case "105"
newclass="111"
newrace="4"
case "106"
newclass="112"
newrace="4"

case "102"
newclass="104"
newrace="4"
case "107"
newclass="111"
newrace="4"
case "108"
newclass="112"
newrace="4"

case "103"
newclass="104"
newrace="4"
case "109"
newclass="111"
newrace="4"
case "110"
newclass="112"
newrace="4"

case "104"
newclass="104"
newrace="4"
case "111"
newclass="111"
newrace="4"
case "112"
newclass="112"
newrace="4"
end select
Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
Else
Response.Write("<center>HATA OLUÞTU.<BR>LÜTFEN BU DURUMU ANASAYFADAN TICKET ATARAK OYUN YÖNETÝCÝSÝNE BÝLDÝRÝNÝZ.</center>")
End If

End If
ElseIf chars("bnation")="2" Then

Set charc=Conne.Execute("select class from userdata where struserid='"&charid&"'")
If Not charc.Eof Then

If charclass="warrior" Then
select case charc("class")
case "201"
newclass="201"
newrace="11"
case "205"
newclass="205"
newrace="11"
case "206"
newclass="206"
newrace="11"

case "202"
newclass="201"
newrace="11"
case "207"
newclass="205"
newrace="11"
case "208"
newclass="206"
newrace="11"

case "203"
newclass="201"
newrace="11"
case "209"
newclass="205"
newrace="11"
case "210"
newclass="206"
newrace="11"

case "204"
newclass="201"
newrace="11"
case "211"
newclass="205"
newrace="11"
case "212"
newclass="206"
newrace="11"
end select

Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
ElseIf charclass="rogue" Then
Select case charc("class")
Case "201"
newclass="202"
newrace="12"
case "205"
newclass="207"
newrace="12"
case "206"
newclass="208"
newrace="12"

case "202"
newclass="202"
newrace="12"
case "207"
newclass="207"
newrace="12"
case "208"
newclass="208"
newrace="12"

case "203"
newclass="202"
newrace="12"
case "209"
newclass="207"
newrace="12"
case "210"
newclass="208"
newrace="12"

case "204"
newclass="202"
newrace="12"
case "211"
newclass="207"
newrace="12"
case "212"
newclass="208"
newrace="12"
end select

Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
ElseIf charclass="mage" Then
select case charc("class")
case "201"
newclass="203"
newrace="13"
case "205"
newclass="209"
newrace="13"
case "206"
newclass="210"
newrace="13"

case "202"
newclass="203"
newrace="13"
case "207"
newclass="209"
newrace="13"
case "208"
newclass="210"
newrace="13"

case "203"
newclass="203"
newrace="13"
case "209"
newclass="209"
newrace="13"
case "210"
newclass="210"
newrace="13"

case "204"
newclass="203"
newrace="13"
case "211"
newclass="209"
newrace="13"
case "212"
newclass="210"
newrace="13"
end select

Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
ElseIf charclass="priest" Then
select case charc("class")
case "201"
newclass="204"
newrace="13"
case "205"
newclass="211"
newrace="13"
case "206"
newclass="212"
newrace="13"

case "202"
newclass="204"
newrace="13"
case "207"
newclass="211"
newrace="13"
case "208"
newclass="212"
newrace="13"

case "203"
newclass="204"
newrace="13"
case "209"
newclass="211"
newrace="13"
case "210"
newclass="212"
newrace="13"

case "204"
newclass="204"
newrace="13"
case "211"
newclass="211"
newrace="13"
case "212"
newclass="212"
newrace="13"
end select

Conne.Execute("Update Userdata Set class='"&newclass&"', race='"&newrace&"', stritem='',strserial='' where struserid='"&charid&"'")
Conne.Execute("Update tb_user set cashpoint=cashpoint-'"&jobcash("ucret")&"' where straccountid='"&accid&"'")
Conne.Execute("insert into logs(ip,islem,islemtarihi) values('"&ips&"','"&charid&" nickli karakter Job deðiþimi yaptý. Yeni Class:"&newclass&", yeni race:"&newrace&" olarak deðiþtirildi.','"&now&"')")
Conne.Execute("exec baslangicitem '"&charid&"'")
Response.Write("<center>Your class has been updated successfully as a "&cla(newclass)&"</center>")
Response.End
Else
Response.Write("<center>HATA OLUÞTU.<BR>LÜTFEN BU DURUMU ANASAYFADAN TICKET ATARAK OYUN YÖNETÝCÝSÝNE BÝLDÝRÝNÝZ.</center>")
End If
Else
Response.Write("hata")
End If

End If
Else
Response.Redirect("powerupstore.asp?change=job")
End If

End If
End If

' PUS ANA BÝTÝÞ
End If
Set pusitem=Conne.Execute("select * from pus_itemleri where type='"&itype&"' order by premiumitem desc,alindi desc,type asc")
Set pusitem2=Conne.Execute("select count(*) toplam from pus_itemleri where type='"&itype&"' ")
Toplam_sonuc =pusitem2("toplam")
git=secur(cint(Request.Querystring("git")))
if git="" or isnumeric(git)=false Then 
git=1
elseif git<=0 Then
git=1
else
End If
ilkkayit=9*(git-1)
if Toplam_sonuc =<ilkkayit Then
Response.Write""
else
pusitem.move(ilkkayit)
	
for i=1 to 9
if pusitem.eof Then exit for
	
if i mod 3=1 Then 
Response.Write "</tr><tr>"
End If %>
                      <td valign="top" width="217"><table background="images/n_table_01.gif" border="0" cellpadding="0" cellspacing="0" height="155" width="213">


                            <tbody><tr><td valign="top">
                              
							<table border="0" style="display:" cellpadding="0" cellspacing="0" width="200">
                                <tbody><tr align="center" valign="middle">
                                  <td class="12_yellow" width="20">&nbsp;</td>

                                  <td colspan="2" style="font-family: 'verdana',Times,serif; color: rgb(0, 0, 0); font-size: 10px; padding-top: 10px;" height="30"><strong>* <%=pusitem("itemismi")%>                             </strong></td>
                                </tr>
                                <tr>
                                  <td align="right" width="78">&nbsp;</td>
                                  <td align="right" width="78"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                                      <tbody><tr>
                                        <td align="right" height="80" valign="bottom"><table border="0" cellpadding="0" cellspacing="0" height="72" width="72">
                                            <tbody><tr>

                                              <td align="center" background="images/table1_01.gif"><a href="javascript:detail_go(<%=pusitem("itemkodu")%>);"><img src="Item/<%=resim(pusitem("resim"))%>" border="0" height="64" width="64"></a></td>
                                          </tr>
                                          </tbody></table></td>
                                      </tr>
                                      <tr>
                                        <td align="center" height="32" valign="bottom"><a href="javascript:detail_go(<%=pusitem("itemkodu")%>);" onMouseOver="MM_swapImage('butt_detail<%=i%>','','images/butt/butt_detail_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt/butt_detail.gif" name="butt_detail<%=i%>" width="72" height="30" border="0" id="butt_detail<%=i%>" /></a></td>
                                    </tr>
                                    </tbody></table></td>
                                  <td align="center" valign="bottom" width="102"><table border="0" cellpadding="0" cellspacing="0" height="108" width="106">

                                      <tbody><tr>
                                        <td align="right"><table border="0" cellpadding="0" cellspacing="0" width="100">
                                            <tbody><tr>
                                              <td style="font-family: 'verdana',Times,serif; color: rgb(0, 0, 0); font-size: 10px;" width="10">-</td>
                                              <td style="font-family: verdana,sans-serif; color: rgb(0, 0, 0); font-size: 11px;" height="16">Ücret :<%=pusitem("ucret")%>
		                           </td>
                                            </tr>
                                            <tr>
                                              <td style="font-family: 'verdana',Times,serif; color: rgb(0, 0, 0); font-size: 10px;">-</td>

                                              <td style="font-family: verdana,sans-serif; color: rgb(0, 0, 0); font-size: 11px;" height="16">Adet <%=pusitem("adet")%> </td>
                                            </tr>
                                            <%if pusitem("kullanimgunu")<>0 Then%><tr>
                                              <td style="font-family: 'verdana',Times,serif; color: rgb(0, 0, 0); font-size: 10px;">-</td>
                                              <td style="font-family: verdana,sans-serif; color: rgb(0, 0, 0); font-size: 11px;" height="16">Süre :<%=pusitem("kullanimgunu")&" Gün"%>
                                            </tr><%End If%>
                                          </tbody></table></td>
                                      </tr>
                                      <tr>
                                        <td align="center" height="30">
                                      <a href="javascript:buy_go(<%=pusitem("itemkodu")%>);" onMouseOver="MM_swapImage('butt_buy<%=i%>','','images/butt/butt_buy_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt//butt_buy.gif" name="butt_buy<%=i%>" width="89" height="30" border="0" id="butt_buy<%=i%>" /></a></td>

                                      </tr>
                                      <tr>
                                        <td align="center" height="30"><a href="javascript:wish_go(<%=pusitem("itemkodu")%>);" onMouseOver="MM_swapImage('butt_cart<%=i%>','','images/butt/butt_cart_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt/butt_cart.gif" name="butt_cart<%=i%>" width="89" height="30" border="0" id="butt_cart<%=i%>" /></a><a href="#"></a></td>
                                    </tr>
                                    </tbody></table></td>
                                </tr>
                              </tbody></table>
										
						
                            </td>

                          </tr>
                        </tbody></table></td>

    <%pusitem.MoveNext
	 next
	 End If %>
                          </tr>
                    </tbody>
                    </table></td>
                </tr>
              <tr>
              <td> <% if not itype="" Then %>
          <table border="0" align="center" cellpadding="0" cellspacing="0">
                    <tbody>
                      <tr>
                        <td height="3"></td>
                        <td height="3"></td>
                        <td height="3"></td>
                      </tr>
                      <tr>
                        <td width="35">
<%dim Toplam
Toplam=cint(Toplam_sonuc/9)
if toplam_sonuc mod 9>0 Then
toplam=toplam+1
End If
if git>1 Then %><a href="powerupstore.asp?git=<%=git-1%>&type=<%=itype%>"><img src="images/butt/butt_arrow_l.gif" border="0" height="17" width="32"></a>
                          <% else %>
                          <img src="images/butt/butt_arrow_l.gif" border="0" height="17" width="32">
                          <%End If%></td>
                        <td align="center"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tbody>
                              <tr>
                                <td width="13"><img src="images/page_01.gif" height="21" width="13"></td>
                                <td class="12_yellow2_140" align="center" background="images/page_02.gif" width="122"><strong><strong>
        <%Dim y
		For y=1 to Toplam
		If y=cint(git) Then
		Response.Write "<font color='#ffee2f' face='verdana' size='2'>"&y&"</font>&nbsp;"
		Else
		Response.Write "<font color='#c3cccc' face='verdana' size='2'><a href=powerupstore.asp?git="&y&"&type="&itype&">"&y&"</a></font>&nbsp;"
		End If
		next %>
                                </strong></strong></td>
                                <td width="15"><img src="images/page_03.gif" height="21" width="13"></td>
                              </tr>
                            </tbody>
                        </table></td>
                        <td align="right" width="35">
			<% 
if cint(git) < cint(toplam) Then %>
                            <a href=powerupstore.asp?git=<%=git+1%>&type=<%=itype%>><img src="images/butt/butt_arrow_r.gif" border="0" height="17" width="32"></a>
                            <% else %>
                            <img src="images/butt/butt_arrow_r.gif" border="0" height="17" width="32">
                            <%End If %></td>
                      </tr>
                    </tbody>
                  </table>
                <% ELSE
			End If %>
              </td></tr>
      </tbody></table></td></tr>
          
            </tbody></table>
</td>
            </tr>
            </tbody>
            </table>
           
</div>
<div style="position: absolute; top: 120px; left: 720px;">

<iframe src="hit.asp?<%=Request.ServerVariables("QUERY_STRING")%>" name="right" allowtransparency="true" background-color="transparent" scrolling="no" width="288" frameborder="0" height="591"></iframe>
</div>
</body></html>
	<%else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>