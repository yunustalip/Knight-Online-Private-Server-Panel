<!--#include file="_inc/conn.asp"-->
<!--#include file="Guvenlik.asp"-->
<!--#include file="function.asp"-->
<!--#include file="md5.asp"-->
<%a=timer
response.expires=0
Response.Charset = "iso-8859-9"
If Not Request.ServerVariables("Script_Name")="/404.asp" Then
yn("/Karakter-Detay")
End If

REFERER_URL = Request.ServerVariables("HTTP_REFERER")
s=Request.ServerVariables("Script_Name")
If InStr(8, REFERER_URL, "/") = 0 Then
REFERER_DOMAIN = REFERER_URL
Else
REFERER_DOMAIN = Left(REFERER_URL, InStr(8, REFERER_URL, "/")-1)
End If


If REFERER_DOMAIN="http://"&Request.ServerVariables("server_name") or  REFERER_DOMAIN="http://www."&Request.ServerVariables("server_name") Then
Else
yn("/Karakter-Detay")
End If
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" >
<style type="text/css">

.adt{
	color:#FFFFFF;
	font-size:11px;
	font-family:Arial, Verdana, Helvetica, sans-serif;
	font-weight:bold
	}
	body,th {
	color: #FFFFFF;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
 
td {
	color: #000000;
	font-family:Verdana, Arial, Helvetica, sans-serif;
	font-size:10px;
}
.style4 {
	color: #5184BD;
	font-weight: bold;
}
.adt{
height:14px
}
</style>
</head>
<body bgcolor="#F9EED8" oncontextmenu="return false"  onselectstart="return false" ondragstart="return false">
<center><br>
<img src="/imgs/karakterdetay.gif"><%Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='KarakterDetay' or PId='Inventory'")
If MenuAyar("PSt")=1 Then
	
Dim lnk,linkp,sid,chard,clang,clanid,ksizsiralama,sira,smbl,sembol,symbol,onl,np1,np2,donatenp,saat,dakika,aut,totalchar,banka,skill,acc1,acc,premium,ia,skillno,sure,skillname,skillnm,king,job,capetop,yuk,gen,capetop2,cape,capem,pelerinm,pelerin,yuks,charid,itemler,dtype,speed,kinds,delay,kind,renk,atack,weight,durability,duration,defans,dodging,incap,daggerac,swordac,clubac,axeac,spearac,bowac,firedam,icedam,ligthdam,posdam,hpdrain,mpdamage,mpdrain,mirrordam,strbon,canbonus,hpbon,dexbon,intbon,mpbon,magicbon,fireres,glares,lightres,magicres,posres,curseres,ReqStr,reqhp,reqdex,Reqint,Reqcha,drtn,iname,itemname2,itemname,slot,maceac,sf,xx,x,skill1,skill2,skill3,skill4,point,clas,skillname1,skillname2,skillname3,zoneid,px,pz,map,pxx,pzz,cleft,ctop

	If Session("login")="ok" Then
	lnk=secur(Trim(Session("sayfa")))
	linkp=Split(lnk,"/")
	If Ubound(linkp)>3 Then
	sid=Trim(linkp(4))
	End If
If Not Len(sid)>0 Then
Response.Write("<br><br><div style=""color:#ff0000;font-weight:bold"">Karakter Bulunamadı</div>")
Response.End
End If
Set chard = Conne.Execute("Select u.Nation, u.Race, u.Class, u.Rank, u.Level,u.Gold, u.Exp, u.Knights, u.Fame, u.Authority, u.Zone, u.PX, u.PZ, u.strSkill, u.loyalty, u.LoyaltyMonthly, u.CreateTime, u.UpdateTime, u.OnlineTime, u.GunlukNp1, u.GunlukNp2,u.HaftalikNp1, u.HaftalikNp2, u.strong, u.sta, u.dex, u.intel, u.cha,u.points, z.bz From USERDATA u, ZONE_INFO z where u.strUserId='"&sid&"' and z.zoneno=u.Zone ")
if not chard.eof Then
if chard("knights")<>0 Then
Set clang =Conne.Execute("Select IDName,scape From KNIGHTS Where IDNum = "&chard("Knights")&"")
clanid=clang("idname")
End If
%>

<style type="text/css">
<!--
.style288 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<script type="text/javascript">
function loadskill(skillbarno){
$.ajax({
   type: 'GET',
   data:'sid=<%=sid%>&skillbarno='+skillbarno,
   url: 'Skills.asp',
   success: function(ajaxCevap) {
      $('div#skills').html(ajaxCevap);
   }
});
}

chngtitle('<%=sid%> > Karakter Detay');
$invenclose = false;
$skillpoint = false;
$userbilgi = false;
function kontrol(olay)
{
	olay = olay || event;

if(olay.keyCode==75){
if ($skillpoint == false){
$('#skillpoint').stop().animate({width: '0%',opacity: 0},1000);
$skillpoint = true;
}
else{
$('#skillpoint').stop().animate({width:'97%',opacity: 1},1000);
$('#skillpoint').css("width","");
$skillpoint = false;
}
}

if(olay.keyCode==85){
if ($userbilgi == false){
$('#userbilgi').stop().animate({width: '0%',opacity: 0},1000);
$userbilgi = true;
}
else{
$('#userbilgi').stop().animate({width:'97%',opacity: 1},1000);
$('#userbilgi').css("width","");
$userbilgi = false;
}
}

if(olay.keyCode==73){
if ($invenclose == false){
$('#inven').stop().animate({width: '0px',left: '512px',opacity: 0},1000);
$invenclose = true;
}
else{
$('#inven').stop().animate({width:'347px',left:'165px',opacity: 1},1000);
$invenclose = false;
}
}


if(olay.keyCode==112){
loadskill(1);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==113){
loadskill(2);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==114){
loadskill(3);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==115){
loadskill(4);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==116){
loadskill(5);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==117){
loadskill(6);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==118){
loadskill(7);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}
if(olay.keyCode==119){
loadskill(8);
document.onhelp = function() {return(false);}
window.onhelp = function() {return(false);}
return false;
}	
}
document.onkeydown = kontrol;

if (window.document.addEventListener) {
window.document.addEventListener("keydown", avoidInvalidKeyStorkes, false);
} else {
window.document.attachEvent("onkeydown", avoidInvalidKeyStorkes);
document.captureEvents(Event.KEYDOWN);
}

function avoidInvalidKeyStorkes(evtArg) {

var evt = (document.all ? window.event : evtArg);
var isIE = (document.all ? true : false);
var KEYCODE = (document.all ? window.event.keyCode : evtArg.which);

var element = (document.all ? window.event.srcElement : evtArg.target);

switch (KEYCODE) {
case 112: //F1
case 113: //F2
case 114: //F3
case 115: //F4
case 116: //F5
case 117: //F6
case 118: //F7
case 119: //F8
case 120: //F9
case 121: //F10
case 122: //F11
case 123: //F12
case 27: //ESCAPE

if (isIE) {
if (KEYCODE == "112") {
document.onhelp = function() { return (false); }
window.onhelp = function() { return (false); }
}

evt.returnValue = false;
evt.keyCode = 0;
} else {
evt.preventDefault();
evt.stopPropagation();
}
break;
default:
window.status = "Done";
}
}
</script>
<%
dim klisira,ksizsira
set ksira=Conne.Execute("select top 100 struserid from userdata where nation="&chard("nation")&" and authority=1 order by loyalty desc, level desc")
set ksizsiralama=Conne.Execute("select top 100 struserid from userdata where nation="&chard("nation")&" and authority=1 order by loyaltymonthly desc, level desc")

For sira=1 to 100
If not ksira.eof Then
If trim(ksira("struserid"))=trim(sid) Then
klisira=sira
Exit For
End If
ksira.MoveNext
ElseIf ksira.Eof Then
Exit For
End If
Next


For sira=1 to 100
If not ksizsiralama.eof Then
If trim(ksizsiralama("struserid"))=trim(sid) Then
ksizsira=sira
Exit For
End If
ksizsiralama.MoveNext
ElseIf ksizsiralama.Eof Then
exit for
End If
next

if klisira="" and ksizsira<>"" Then
smbl="ksizsembol"
sembol=ksizsira

elseif klisira<>"" and ksizsira<>"" Then

if klisira>ksizsira Then
smbl="ksizsembol"
sembol=ksizsira
else
smbl="klisembol"
sembol=klisira
End If

elseif klisira<>"" and ksizsira="" Then
smbl="klisembol"
sembol=klisira
End If

if smbl="klisembol" Then

if sembol=1 Then
symbol="<img src=""/imgs/1.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>1 and sembol<5 Then
symbol="<img src=""/imgs/2.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>4 and sembol<10 Then
symbol="<img src=""/imgs/3.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>9 and sembol<26 Then
symbol="<img src=""/imgs/4.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>25 and sembol<51 Then
symbol="<img src=""/imgs/5.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>50 and sembol<101 Then
symbol="<img src=""/imgs/006.gif"" align=""absmiddle"">&nbsp;"
else
symbol=""
End If

elseif smbl="ksizsembol" Then
if sembol=1 Then
symbol="<img src=""/imgs/001.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>1 and sembol<5 Then
symbol="<img src=""/imgs/002.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>4 and sembol<10 Then
symbol="<img src=""/imgs/003.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>9 and sembol<26 Then
symbol="<img src=""/imgs/004.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>25 and sembol<51 Then
symbol="<img src=""/imgs/005.gif"" align=""absmiddle"">&nbsp;"
elseif sembol>50 and sembol<101 Then
symbol="<img src=""/imgs/006.gif"" align=""absmiddle"">&nbsp;"
else
symbol=""
End If

End If

set onl=Conne.Execute("select strcharid,np,np2 from currentuser where strcharid='"&sid&"'")
if onl.eof Then
np1=0
np2=0
else
np1=chard("loyalty")-onl("np")
np2=chard("loyaltymonthly")-onl("np2")
End If

Set Kesen=Conne.Execute("Select count(Kesen) As Kesen From Deathlog Where KesenIrk>0 And Kesen='"&Sid&"'")
Set Kesilen=Conne.Execute("Select count(Kesilen) As Kesilen From Deathlog Where Irk>0 And Kesilen='"&Sid&"'")
%>
<table cellpadding="0" cellspacing="0" oncontextmenu="return false" style="margin-left:50px">
    <tr align="center">
    <td width="400" rowspan="2" align="right" valign="top" style="padding-top:30px">
	  <table width="100%" border="0" align="left" cellpadding="2" cellspacing="1" oncontextmenu="return false">
	    <tbody>
	      <tr>
	        <td height="4" colspan="2" align="center" valign="top" bgcolor="#CC0000" class="style288">Karakter Bilgisi</td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Karakter</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%=sid%></td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Clan</strong></td>
            <td align="center" width="55%" height="4">
              <%if clanid<>"" Then%>
              <a href="Clan-Detay/,<%=chard("knights")%>" onClick="pageload('Clan-Detay/,<%=chard("knights")%>');return false" class="link1"><%Response.Write clang("idname")
	End If%></a>	</td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Np</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%Response.Write ayir(chard("loyalty"))&" / "&ayir(chard("loyaltymonthly"))%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Haftalık Np</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B"><%Response.Write ayir(chard("HaftalikNp1")+(np1))&" / "&ayir(chard("HaftalikNp2")+(np2))%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Günlük Np</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%Response.Write ayir(chard("gunluknp1")+(np1))&" / "&ayir(chard("gunluknp2")+(np2))%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Clana Bağışlanan Np</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B"><%set donatenp=Conne.Execute("select np from npdonate where userid='"&sid&"' and clan="&chard("knights"))
if not donatenp.eof Then
Response.Write donatenp("np")
else
Response.Write "0"
End If%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Kesilen Karakter Sayısı</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%=Kesen("Kesen")%></td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Ölme Sayısı</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B"><%=Kesilen("Kesilen")%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Tür</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%=cla(chard("Class"))%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Irk</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B"><% if chard("Nation")="1" Then
		  Response.Write "<img src='/imgs/karuslogo.gif' />"
          elseif chard("Nation")="2" Then
		  Response.Write "<img src='/imgs/elmologo.gif' />"
          End If %></td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Durum</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B">
              <%if not onl.eof Then
	Response.Write "<font color=""#ff0000""><b>Oyunda !</b></font>"
	else
	Response.Write "<font color=""#666666""><b>Çevrimdışı</b></font>"
	End If%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Toplam Online Süresi</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B">
            <%

if chard("onlinetime")>60 Then
dakika=int(chard("onlinetime")/60)
else
saat=0
dakika=0
End If
if dakika>60 Then
saat=int(dakika/60)
dakika=int(dakika mod 60)
else

End If


if saat>0 Then Response.Write saat&" Saat "
if dakika>0 Then Response.Write dakika&" Dk. "%>        </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Karakter Seviyesi</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%aut=chard("authority")
	if aut="1" Then
	Response.Write "Oyuncu"
	elseif  aut="0" Then
	Response.Write "<font color=""red""><b>Game Master</b></font>"
	elseif aut="255" Then
	Response.Write "Banlanmış Oyuncu"
	elseif aut="2" or aut="11" Then
	Response.Write "Muteli Oyuncu"
	End If
	if chard("rank")="1" Then
	Response.Write " & Kral"
	End If%></td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Oluşturulma Tarihi</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B"><%=chard("Createtime")%></td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Son Giriş Tarihi</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%=chard("updatetime")%></td>
          </tr>
	      <tr>
	        <td height="4" colspan="2" align="center" bgcolor="#CC0000" class="style288">Hesap Bilgisi</td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="color:#89640B"><strong>Toplam Karakter</strong></td>
            <td align="center" width="55%" height="4" style="color:#89640B"><%
	  Set totalchar = Conne.Execute("Select straccountid,bCharNum From Account_Char where strcharid1 = '"&sid&"' or strcharid2 = '"&sid&"' or strcharid3 = '"&sid&"'")
	  if not totalchar.eof Then
	   Response.Write totalchar("bCharNum")
	else
	Response.Write "0"
	End If%>          </td>
          </tr>
	      <tr>
	        <td width="45%" height="4" align="left" style="background-color:#DCD1BA;color:#89640B"><strong>Bankadaki Para</strong></td>
            <td align="center" width="55%" height="4" style="background-color:#DCD1BA;color:#89640B"><%
if not totalchar.eof Then
Set banka =Conne.Execute("Select nmoney From warehouse Where strAccountid = '"&totalchar("straccountid")&"'")
if not banka.eof Then
Response.Write banka("nMoney")
else
Response.Write "0"
End If
banka.close
set banka=nothing
else
Response.Write "0"
End If%></td>
          </tr>
	      <tr>
	        <td></td>
          </tr>
	      <% if Session("yetki")="1" Then
		if aut=255 Then%>
	      <tr>
	        <td align="center" ><input name="button" type="button" class="styleform" style=" border-style:solid; border-color:#FFFFFF; width: 150px" onClick="gpopup('<%="GmPage/Gamem.asp?user=bankaldir&nick="&sid%>')"  value="Ban Kaldır" />          </td>
  </tr><%else
if not onl.eof Then%>
	      <tr>
	        <td align="center" ><input name="button" type="button" class="styleform" style=" border-style:solid; border-color:#FFFFFF; width:150px" onClick="gpopup('GmPage/Gamem.asp?user=ban&nick=<%=sid%>')"  value="Banla" />         </td>
            <td align="center" ><input name="button" type="button" class="styleform" style=" border-style:solid; border-color:#FFFFFF; width:185px" onClick="gpopup('GmPage/Gamem.asp?user=dc&nick=<%=sid%>')"  value="Oyundan At (Disconnect)" />         </td>
        </tr>
	      <%else%>
	      <tr>
	        <td align="center" ><input name="button" type="button" class="styleform" style=" border-style:solid; border-color:#FFFFFF; width:150px" onClick="gpopup('GmPage/gamem.asp?user=ban&nick=<%=sid%>')"  value="Banla" />         </td>
        </tr>
  <%End If
End If
		End If%>
	      </tbody>
        </table></td>
    <td align="right" valign="top"><div align="right" style="padding-right:30px;height:70px">
  <%Set skill=Conne.Execute("select * from USER_SAVED_MAGIC where strcharid='"&sid&"'")
If Not totalchar.Eof Then
Set acc=Conne.Execute("select premiumtype,premiumdays from tb_user where straccountid='"&totalchar("straccountid")&"'")

If Not acc.Eof Then
If acc("premiumtype")=1 and round(acc("premiumdays")-now)>0 Then
premium="Premium Service "&round(acc("premiumdays")-now)+1&" Days"
Else
premium="Premium Service 0 Day&nbsp;&nbsp;&nbsp;&nbsp;"
End If
Response.Write "<span style=""color:fff;font-weight:bold;position:relative;top:18px;right:5px"">"&premium&"</span><br><img src=""/imgs/premium.gif""><br>"
End If
End If
Response.Write "<img src=""/imgs/skill.gif"">"
If Not skill.Eof Then
For ia=1 To 10
skillno=skill("nskill"&ia)
sure=skill("nduring"&ia)
If skillno<>0 Then
Set skillname=Conne.Execute("select enname,krname from magic where magicnum="&skillno)
If Not skillname.eof Then

if instr(skillname(1),"?")>0 Then
skillnm=skillname(0)
else
skillnm=skillname(1)
End If

End If
Response.Write "<img src=""/skill/skillicon_"&mid(skillno,5,2)&"_"&mid(skillno,1,4)&".bmp"" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=white>"&server.htmlencode(skillnm)&("<br>Kalan Süre: "&sure&" Dakika")&"', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"

skillnm="Unknown Skill"
End If
Next
End If

Response.Write "<br>"
%>
    </div></td>
    </tr>
  <tr align="center">
    <td width="500" align="center">
	<%if chard("rank")="1" Then
	king=1
	else
	king=0
	End If
	if king="1" Then	
	Response.Write "<img src=""/imgs/king.gif"" width=""50"" style=""position:relative;top:0px;z-index:5""><br>"
	End If
	
	If clanid<>"" Then
	Response.Write "<font color='#E90621' style='font-size:11px'><b>"&clang("idname")&"</b></font><br>"
	End If
	
	Response.Write("<table cellpadding=""0"" cellspacing=""0""><tr><td valign=""middle"">"&symbol&"</td><td valign=""middle""><span class=""style4""><font style=""font-size:11px"">"&sid&"</span></td><tr>")
	If chard("fame")="1"  Then
	Response.Write "<tr><td></td><td valign=""top""><img src=""/imgs/cleader.gif"" width=""100%"" height=""3""></td></tr>"
	End If 
	Response.Write "</table>"
        If chard("race")="1" Then
	job="<img src='/imgs/karusbuyukwarrior.gif' align=""middle"" style=""position:relative;z-index:3"">"
	capetop=284
	yuk=220
	gen=115
	alttop=40
	Elseif chard("race")="2" Then
	job="<img src='/imgs/karuserkek.gif' align=""middle"" style=""position:relative;z-index:3"">"
	capetop=233
	yuk=190
	gen=100
	alttop=65
	Elseif chard("race")="3" Then
	job="<img src='/imgs/karuscucemage.gif' align=""middle"" style=""position:relative;z-index:3"">"
	capetop=185
	yuk=195
	gen=100
	alttop=60
	Elseif chard("race")="4" Then
	job="<img src='/imgs/karuskadin.gif' align=""middle"" style=""position:relative;z-index:3"">"
	capetop=220
	yuk=190
	gen=65
	cleft=-2
	alttop=80
	ElseIf chard("race")="11" Then
	job="<img src=""/imgs/humanwarrior.gif"" align=""top"" style=""position:relative;z-index:3"" height=""325"">"
	capetop=260
	yuk=200
	gen=120
	cleft=-8
	alttop=50
	Elseif chard("race")="12" Then
	job="<img src='/imgs/humanerkek.gif' align=""middle"" style=""position:relative;z-index:3"" height=""325"">"
	capetop=274
	yuk=215
	gen=110
	alttop=50
	Elseif chard("race")="13" Then
	job="<img src=""/imgs/humankadin.gif"" align=""middle"" style=""position:relative;z-index:3"" height=""325"">"
	capetop=284
	yuk=220
	gen=90
	alttop=50
	Else
	job=""
	End If 

Response.Write job
if clanid<>"" Then
cape=clang("scape")
if king="1" Then
cape="99"
capem=""
End If
if not cape=-1 Then
	if len(cape)=3 Then
	capem=left(cape,1)
	pelerinm="../imgs/cape/cloak_m_0"&capem&".gif"
	cape=mid(cape,2,2)
	End If
	if len(cape)=1 Then
	cape="0"&cape
	End If
pelerin="../imgs/cape/cloak_c_"&cape&".gif"

If cape<>"" Then
Response.Write "<br><img src="""&pelerin&""" height="""&yuk&""" width="""&gen&""" style=""position:relative;top:-"&capetop&"px;left:"&cleft&"px;z-index:1"">"
If capem<>"" Then
Response.Write "<br><img src="""&pelerinm&""" height="""&yuk&""" width="""&gen&"""  style=""position:relative;top:-"&capetop+yuk&"px;left:"&cleft&"px;z-index:2"">"
End If
End If

Else
cape=""
End If
End If
If aut=255 Then
Response.Write "<br><img src=""imgs/banned.png"" width=""150"" style=""position:relative;"
If cape<>"" And capem<>"" Then
Response.Write "top:-"&capetop+yuk-alttop+250&"px"
ElseIf cape<>"" and capem="" Then
Response.Write "top:-"&capetop-alttop+250&"px"
Else
Response.Write "top:-"&capetop-yuk-alttop+250&"px"
End If
Response.Write ";z-index:6"">"
capetop=capetop+75
Elseif aut=11 Then
Response.Write "<br><img src=""imgs/mute.gif"" width=""150"" style=""position:relative;"
If cape<>"" And capem<>"" Then
Response.Write "top:-"&capetop+yuk-alttop+250&"px"
ElseIf cape<>"" and capem="" Then
Response.Write "top:-"&capetop-alttop+250&"px"
Else
Response.Write "top:-"&capetop-yuk-alttop+250&"px"
End If
Response.Write ";z-index:6"">"
capetop=capetop+75
End If%></td>
  </tr>
</table>
</td></tr></table>
<div id="alt" style="position:relative;<%If cape<>"" And capem<>"" Then
Response.Write "top:-"&capetop+yuk-alttop&"px"
ElseIf cape<>"" and capem="" Then
Response.Write "top:-"&capetop-alttop&"px"
Else
Response.Write "top:-"&capetop-yuk-alttop&"px"
End If
%>">
<br>
<div align="left" style="padding-left:30px;position:relative;width:700px;height:60px" id="skills" oncontextmenu="return false" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%skillbarno=1
function clss2(tur)
select case tur
case "101", "105", "106", "201", "205", "206"
clss2="Warrior"
case "102", "107", "108", "202", "207", "208"
clss2="Rogue"
case "103", "109", "110", "203", "209", "210"
clss2="Mage"
case "104", "111", "112", "204", "211", "212"
clss2="Priest"
Case else
clss2="Unknown"
end select
end function


Conne.Execute("exec skillbardecode '"&sid&"'")
dim snm,skillnum,skiln,des,sklnm,slvl
for sir=1 to 8
skillnm="No Skill"
set snm=Conne.Execute("select skillno from skillbar where charid='"&sid&"' and sira="&sir&" and satir="&skillbarno&"")
if not snm.eof Then
skillnum=snm("skillno")

set skillname=Conne.Execute("select enname,krname,description from magic where magicnum="&skillnum)
if not skillname.eof Then
skiln=mid(skillnum,4,1)
des=skillname("description")
if instr(skillname(1),"?")>0 Then
skillnm=skillname(0)
else
skillnm=skillname(1)
End If

End If

if clss2(chard("Class"))="Warrior" Then
if skiln=5 Then
sklnm="Attack"
elseif skiln=6 Then
sklnm="Defense"
elseif skiln=7 Then
sklnm="Berserker"
elseif skiln=8 Then
sklnm="Master"
End If
End If

if clss2(chard("Class"))="Mage" Then
if skiln=5 Then
sklnm = "Flame"
elseif skiln=6 Then
sklnm = "Glacier"
elseif skiln=7 Then
sklnm = "Lightning"
elseif skiln=8 Then
sklnm="Master"
End If
End If

if clss2(chard("Class"))="Priest" Then
if skiln=5 Then
sklnm = "Heal"
elseif skiln=6 Then
sklnm = "Buff"
elseif skiln=7 Then
sklnm = "Debuff"
elseif skiln=8 Then
sklnm="Master"
End If
End If

if clss2(chard("Class"))="Rogue" Then
if skiln=5 Then
sklnm = "Archer"
elseif skiln=6 Then
sklnm = "Assassin"
elseif skiln=7 Then
sklnm = "Explore"
elseif skiln=8 Then
sklnm="Master"
End If
End If

slvl="("&sklnm&" "&mid(skillnum,5,2)&")"

if mid(skillnum,4,1)="0" Then
slvl="(item)"
End If

Response.Write "<span style=""position:relative;margin-left:2px;z-index:2""  onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=white>"&server.htmlencode(skillnm)&slvl&"<br><br>"&des&"', ABOVE, WIDTH, 300,CELLPAD, 5, 5, 5);"" onMouseOut=""return nd();""><img src=""skill/skillicon_"&mid(skillnum,5,2)&"_"&mid(skillnum,1,4)&".bmp""></span>"
else
Response.Write "<span style=""position:relative;margin-left:2px;z-index:2"" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=white>No Skill', ABOVE, WIDTH, 200,CELLPAD, 0, 0, 5);"" onMouseOut=""return nd();"")""><img src=""skill/skillicon_enigma.bmp""></span>"
End If
des=""
Next
Conne.Execute("Delete skillbar where charid='"&sid&"'")

%><img src="../imgs/skillbar.gif" style="position:relative; left:-303px; top: 3px; z-index:1">
<div style="position:relative;left: 13px; top:-33px; width:10px;z-index:3"><img src="imgs/yukok.gif" border="0" width="9" height="9" style="position:relative;z-index:3"><br><font style="color:#FFFFFF;font-size:10px;position:relative; z-index:3"><b>1</b></font><a onClick="loadskill('2');"><br><img src="imgs/asok.gif" border="0" width="9" height="9" style="position: relative; top: -0px; width: 9px; height: 10px; z-index: 10;"></a><br>
<img src="imgs/skillbartop.gif" style="position:relative;top:-33px;left:13px;z-index:5"></div>
</div><br>

<div align="left" id="userbilgi" style="position:relative;padding-left:30px;height: 400px;">
<script type="text/javascript">
function kulbilgi(sid,detay){
$.ajax({
   url: 'userbilgi/'+sid+'/'+detay+'.html',
   success: function(ajaxCevap) {
      $('div#userbilgi').html(ajaxCevap);
   }
});
}
</script><script language="JavaScript" type="text/JavaScript">
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

<style type="text/css">
.style6 {
	color: #FFFFFF;
	font-weight: bold;
}
</style><table width="100%" oncontextmenu="return false"><tr><td>
<%If detay="" Then
detay="karakter"
End If

If detay="karakter" Then

Set clang =Conne.Execute("Select IDNum,IDName From KNIGHTS Where IDNum = "&chard("Knights")&"")
Set lvlup = Conne.Execute("Select Level,Exp From LEVEL_UP Where Level = "&chard("Level")&"")
%>

<table width="301" height="382" border="0" align="left" cellpadding="0" cellspacing="0" style="background:url(/imgs/statbar.gIf); background-repeat:no-repeat; ">
  <tr>
    <td colspan="3"><div style="position: relative; top: 24px; left: 0px;"><a href="userbilgi/<%=sid%>/karakter" onClick="kulbilgi('<%=sid%>','karakter');return false" onMouseOver="MM_swapImage('char','','/imgs/tabcharon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabchar.gIf" name="char" id="char" border="0"></a></div>
<div style="position: relative; top: -1px; left: 110px;"><a href="userbilgi/<%=sid%>/clan" onClick="kulbilgi('<%=sid%>','clan');return false" onMouseOver="MM_swapImage('clan','','/imgs/tabclanon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabclan.gIf" name="clan" id="clan" border="0"></a></div>
<div style="position: relative; top: -26px; left: 218px;"><a href="userbilgi/<%=sid%>/friend" onClick="kulbilgi('<%=sid%>','friend');return false" onMouseOver="MM_swapImage('friend','','/imgs/tabfriendon.gIf',1)" onMouseOut="MM_swapImgRestore()"><img src="/imgs/tabfriend.gIf" name="friend" id="friend" border="0"></a></div></td>
  </tr>
  <tr>
    <td height="25" colspan="3" class="style6" style="padding-left:15px;padding-top:-15px"><%=sid%></td>
    <td colspan="2" rowspan="2"></td>
    <td width="56" rowspan="2"></td>
  </tr>
  <tr>
    <td height="20" colspan="3" valign="top" class="style6" style="padding-left:15px"><font color="#DFC68C">
      <%=cla2(chard("Class"))%>
    </font></td>
  </tr>
  <tr>
    <td height="18" colspan="2"></td>
    <td width="48" align="center" class="style6" style="padding-left:-55px"><%=chard("Level")%></td>
    <td align="center" class="style6"></td>
    <td colspan="2" align="center" valign="middle" class="style6"><% If chard("Nation")="1" Then
		  Response.Write "Karus"
          elseIf chard("Nation")="2" Then
		  Response.Write "El-Morad"
          End If %></td>
  </tr>
  <tr>
    <td height="32" colspan="2"></td>
    <td height="32" colspan="3" align="center" class="style6"><%=chard("Exp")%> / <%=lvlup("Exp")%></td>
    <td height="32" align="center"></td>
  </tr>
  <tr>
    <td height="23" colspan="2"></td>
    <td height="23" colspan="3" align="center" valign="middle" class="style6"><%=chard("Loyalty")%> / <%=chard("LoyaltyMonthly")%></td>
    <td height="23" align="center"></td>
  </tr>
  <tr>
    <td width="44" valign="bottom">&nbsp;</td>
    <td height="46" colspan="2" align="center" valign="bottom" class="style6"><%=chard("Strong")%></td>
    <td width="55" valign="bottom" class="style6">&nbsp;</td>
    <td width="55" align="center" valign="bottom" class="style6"><%=chard("Cha")%></td>
    <td align="center" valign="bottom">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="48" colspan="2" align="center" valign="middle" class="style6"><%=chard("Sta")%></td>
    <td width="55">&nbsp;</td>
    <td align="center" valign="middle" class="style6"><%=chard("Intel")%></td>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="25" colspan="2" align="center" valign="middle" class="style6"><%=chard("Dex")%></td>
    <td width="55">&nbsp;</td>
    <td height="25" align="center" valign="middle" class="style6"><%=chard("Points")%></td>
    <td height="25" align="center">&nbsp;</td>
  </tr>
    <td height="70"></td>
  </tr>
</table>
<%


End If

%></td></tr>
</table>
</div>
<%Function BinaryToString(Binary)
  Dim cl1, cl2, cl3, pl1, pl2, pl3
  Dim L
  cl1 = 1
  cl2 = 1
  cl3 = 1
  L = LenB(Binary)
  
  Do While cl1<=L
    pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
    cl1 = cl1 + 1
    cl3 = cl3 + 1
    If cl3>300 Then
      pl2 = pl2 & pl3
      pl3 = ""
      cl3 = 1
      cl2 = cl2 + 1
      If cl2>200 Then
        pl1 = pl1 & pl2
        pl2 = ""
        cl2 = 1
      End If
    End If
  Loop
  BinaryToString = pl1 & pl2 & pl3
End Function

skill1=(BinaryToString(midb(chard("strskill"),11,2)))
if skill1<>"" Then
skill1=asc(skill1)
else
skill1="0"
End If
skill2=(BinaryToString(midb(chard("strskill"),13,2)))
if skill2<>"" Then
skill2=asc(skill2)
else
skill2="0"
End If
skill3=(BinaryToString(midb(chard("strskill"),15,2)))
if skill3<>"" Then
skill3=asc(skill3)
else
skill3="0"
End If
skill4=(BinaryToString(midb(chard("strskill"),17,2)))
if skill4<>"" Then
skill4=asc(skill4)
else
skill4="0"
End If
point=(BinaryToString(midb(chard("strskill"),1,2)))
if point<>"" Then
point=asc(point)
else
point="0"
End If
function cla3(tur)

select case tur
case "101", "105", "106", "201", "205", "206"
cla3="Warrior"
case "102", "107", "108", "202", "207", "208"
cla3="Rogue"
case "103", "109", "110", "203", "209", "210"
cla3= "Mage"
case "104", "111", "112", "204", "211", "212"
cla3= "Priest"
Case else
cla3= "Unknown"
end select

end function

clas=cla3(chard("class"))

if clas="Warrior" Then
skillname1="Attack"
skillname2="Defense"
skillname3="Berserker"
elseif clas = "Mage" Then
skillname1 = "Flame"
skillname2 = "Glacier"
skillname3 = "Lightning"
elseif clas = "Priest" Then
skillname1 = "Heal"
skillname2 = "Buff"
skillname3 = "Debuff"
elseif clas = "Rogue" Then
skillname1 = "Archer"
skillname2 = "Assassin"
skillname3 = "Explore"
End If
%>
<style>
.skl{
 color:#e6debb;
 font-family: Arial, Helvetica, sans-serif;
 font-size:10px;
 font-weight:bold;
}
* html .container1 {height: 1%;}
</style>
<div align="left" id="skillpoint" style="padding-left:10px;z-index:1;height:158px">

<table width="347" border="0" oncontextmenu="return false" background="/imgs/skillpoint.gif" style="color:#e6debb; font-family: Arial, Helvetica, sans-serif; font-size:14px">
<tr><td height="57" colspan="2">&nbsp;</td><td width="61" height="57">&nbsp;</td><td width="24" height="57">&nbsp;</td><td width="62" height="57">&nbsp;</td><td width="65" height="57">&nbsp;</td><td width="31" height="57">&nbsp;</td></tr>
<tr><td width="11" height="24">&nbsp;</td>
  <td height="24" colspan="2" align="center" class="skl"><%=cla2(chard("class"))%></td>
  <td height="24">&nbsp;</td><td height="24">&nbsp;</td><td height="24">&nbsp;</td><td height="24">&nbsp;</td></tr>
<tr><td height="20">&nbsp;</td>
  <td width="63" height="20" class="skl"><%=skillname1%></td>
  <td height="20" class="skl"><%=skill1%></td><td height="20">&nbsp;</td><td height="20" class="skl"><%=skillname2%></td>
  <td height="20" class="skl"><%=skill2%></td><td height="20">&nbsp;</td></tr>
<tr><td height="22">&nbsp;</td>
  <td height="22" class="skl"><%=skillname3%></td>
  <td height="22" class="skl"><%=skill3%></td><td height="22">&nbsp;</td><td height="22" class="skl">Master</td>
  <td height="22" class="skl"><%=skill4%></td><td height="22">&nbsp;</td></tr>
  <tr><td height="20">&nbsp;</td>
  <td width="63" height="20">&nbsp;</td>
  <td height="20">&nbsp;</td><td height="20">&nbsp;</td><td height="20" class="skl">Points</td>
  <td height="20" class="skl"><%=point%></td><td height="20">&nbsp;</td></tr>
</table>
</div>
<%if isnumeric(zoneid)=false or isnumeric(px)=false or isnumeric(pz)=false Then
Response.End
End If
zoneid=chard("zone")
if zoneid="1" or zoneid="2" or zoneid="21" or zoneid="201" or zoneid="202" or zoneid="11" or zoneid="12" Then

if zoneid=11 or zoneid=12 Then
map=1112
else
map=zoneid
End If

pxx=chard("px")
pzz=chard("pz")

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
%>
<style type="text/css">
.bdy {
background-image:url(/imgs/Maps/<%Response.Write map&".jpg"%>);
background-position:top left;
background-repeat:no-repeat;
margin-left:0px;
margin-top:0px;

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
<div id="karakterzoneid" align="center" oncontextmenu="return false">
<div style="font-weight:bold"><%Response.Write chard("bz")&" ("&px&","&pz&")"%></div>
<div class="bdy" style="position:relative;width:512px;height:512px">
<%
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
Response.Write "<img src=""imgs/Red_Arrow_Down.gif"" width=""15"" height=""20"" style=""position:relative;left:"&cleft-255&"px; top:"&ctop-20&"px;z-index:8;"">"&vbcrlf

Else
Response.Write "Bu Haritada Bulunmamaktadır."
End If%>
</div>

</div>

<div align="left" id="inven" style="position:relative;width:347px<%MenuAyar.MoveNext
If MenuAyar("PSt")=1 Then Response.Write ";background:url(/Item/inventory_bg.JPG)"%>;top:-1100px;left:165px;height:576px;background-repeat:no-repeat;z-index:2" oncontextmenu="return false">
<%Dim items(42)
Dim adet(42)
If MenuAyar("PSt")=1 Then 
charid=sid
Conne.Execute("delete INVENTORY where struserid='"&charid&"'")
Conne.Execute("exec item_decode3 '"&charid&"'")
Set Itemler=Conne.Execute("SELECT i.itemtype,i.delay,i.kind,i.strname,i.damage,i.weight,i.duration,i.ac,i.Evasionrate,i.Hitrate,i.daggerac,i.swordac,i.axeac,i.spearac,i.bowac,i.maceac,i.firedamage,i.icedamage,i.LightningDamage,i.poisondamage,i.hpdrain,i.MPDamage,i.mpdrain,i.MirrorDamage,i.StrB,i.StaB,i.MaxHpB,i.DexB,i.IntelB,i.MaxMpB,i.ChaB,i.FireR,i.coldr,i.LightningR,i.magicr,i.poisonr,i.curser,i.reqstr,i.ReqSta,i.reqdex,i.reqintel,i.reqcha,i.Countable,inv.num,inv.stacksize,inv.durability,inv.struserid,inv.inventoryslot FROM INVENTORY inv, item i WHERE inv.struserid='"&charid&"' and inv.num=i.num  group by i.itemtype,i.delay,i.kind,i.strname,i.damage,i.weight,i.duration,i.ac,i.Evasionrate,i.Hitrate,i.daggerac,i.swordac,i.axeac,i.spearac,i.bowac,i.maceac,i.firedamage,i.icedamage,i.LightningDamage,i.poisondamage,i.hpdrain,i.MPDamage,i.mpdrain,i.MirrorDamage,i.StrB,i.StaB,i.MaxHpB,i.DexB,i.IntelB,i.MaxMpB,i.ChaB,i.FireR,i.coldr,i.LightningR,i.magicr,i.poisonr,i.curser,i.reqstr,i.ReqSta,i.reqdex,i.reqintel,i.reqcha,i.Countable,inv.num,inv.stacksize,inv.durability,inv.struserid,inv.inventoryslot order by inv.inventoryslot asc")


If Not itemler.Eof Then
Do While not itemler.Eof

Dtype=Itemler("ItemType")
Speed=Itemler("delay")
Kinds=Itemler("kind")

If speed>0 and speed<90 Then
delay = "Atack Speed : Very Fast<br>"
ElseIf speed>89 and speed<111 and not kinds=>91 and not kinds=<95 Then
Delay = "Atack Speed : Fast<br>"
ElseIf speed>110 and speed<131 Then
delay = "Atack Speed : Normal<br>"
elseIf speed>130 and speed<151 Then
Delay = "Atack Speed : Slow<br>"
elseIf speed>150 and speed<201 Then
delay = "Atack Speed : Very Slow<br>"
Else
delay=""
End If

if kinds=11 Then
kind="Dagger"
elseif kinds =21 Then
kind="One-handed Sword"
elseif kinds = 22 Then
kind="Two-handed Sword"
elseif kinds =31 Then
kind= "Axe"
elseif kinds = 32 Then
kind="Two-handed Axe"
elseif kinds = 41 Then
kind="Club"
elseif kinds = 42 Then
kind="Two-handed Club"
elseif kinds = 51 Then
kind="Spear"
elseif kinds = 52 Then
kind="Long Spear"
elseif kinds = 60 Then
kind="Shield"
elseif kinds = 70 Then
kind="Bow"
elseif kinds = 71 Then
kind="Crossbow"
elseif kinds = 91 Then
kind="Earring"
elseif kinds = 92 Then
kind="Necklace"
elseif kinds = 93 Then
kind="Ring"
elseif kinds = 94 Then
kind="Belt"
elseif kinds = 95 Then
kind="Lune Item"
elseif kinds = 98 Then
kind="Upgrade Scroll"
elseif kinds = 110 Then
kind="Staff"
elseif kinds = 120 Then
kind="Staff"
elseif kinds = 210 Then
kind="Warrior Armor"
elseif kinds = 220 Then

kind="Rogue Armor"
elseif kinds = 230 Then
kind="Magician Armor"
elseif kinds = 240 Then
kind="Priest Armor"
else
kind=""
End If

if dtype=0 Then
if dtype=0 and kinds=255 Then
dtype="(Scroll)"
renk="white"
else
dtype="(Non Upgrade Item)"
renk="white"
End If
elseif dtype=1 Then
dtype="(Magic Item)"
renk="blue"
elseif dtype=2 Then
dtype="(Rare Item)"
renk="yellow"
elseif dtype=3 Then
dtype="(Craft Item)"
renk="lime"
elseif dtype=4 Then
dtype="(Unique Item)"
renk="#DFC68C"
elseif dtype=5 Then
dtype="(Upgrade Item)"
renk="#CE8DC5"
elseif dtype=6 Then
dtype="(Event Item)"
renk="lime"
End If

if itemler("Damage")>0 Then 
atack="Attack Power : "&itemler("Damage") & "<br>"
else
atack=""
End If
if itemler("Weight")>1 Then 
if len(itemler("Weight"))=2 Then
sf="00"
elseif len(itemler("weight"))=3 Then
sf="0"
End If
weight="Weight : "&left(itemler("Weight"),2)&"."&mid(itemler("weight"),3,1)&sf&"<br>"
else
weight=""
End If
if itemler("durability")>1 and not kinds=255 Then
Durability="Current Durability : "&itemler("durability")&"<br>"
else
Durability=""
End If
if itemler("Duration")>1 and itemler("ItemType")=0 Then
duration="Quantity : "&itemler("Duration") & "<br>"
elseif itemler("Duration")>1  Then 
duration="Max Durability : "&itemler("Duration") & "<br>"
else
duration=""
End If
if itemler("Ac")>0 Then 
defans="Defense Ability : "&itemler("Ac") & "<br>"
else
defans=""
End If
if itemler("Evasionrate")>0 Then
dodging="Increase Dodging Power by : "&itemler("Evasionrate")&"<br>"
else
dodging=""
End If
if itemler("Hitrate")>0 Then
incap="Increase Attack Power by  : "&itemler("Hitrate")&"<br>"
else
incap=""
End If


if itemler("DaggerAc")>0 Then 
daggerac="Defense Ability (Dagger) : "&itemler("DaggerAc") & "<br>"
else
daggerac=""
End If
if itemler("SwordAc")>0 Then 
swordac="Defense Ability (Sword) : "&itemler("SwordAc") & "<br>"
else
swordac=""
End If
if itemler("MaceAc")>0 Then 
clubac="Defense Ability (Club) : "&itemler("MaceAc") & "<br>"
else
clubac=""
End If
if itemler("AxeAc")>0 Then 
axeac="Defense Ability (Axe) : "&itemler("AxeAc") & "<br>"
else
axeac=""
End If
if itemler("SpearAc")>0 Then 
spearac="Defense Ability (Spear) : "&itemler("SpearAc") & "<br>"
else
spearac=""
End If
if itemler("BowAc")>0  Then
bowac="Defense Ability (Arrow) : "&itemler("BowAc") & "<br>"
else
bowac=""
End If
if itemler("FireDamage")>0  Then
firedam="Flame Damage : "&itemler("FireDamage") & "<br>"
else
firedam=""
End If
if itemler("IceDamage")>0  Then
icedam="Glacier Damage : "&itemler("IceDamage") & "<br>"
else
icedam=""
End If
if itemler("LightningDamage")>0  Then
ligthdam="Lightning Damage : "&itemler("LightningDamage") & "<br>"
else
ligthdam=""
End If
if itemler("PoisonDamage")>0  Then
posdam="Poison Damage : "&itemler("PoisonDamage") & "<br>"
else
posdam=""
End If
if itemler("HPDrain")>0 Then
hpdrain="HP Recovery : "&itemler("HPDrain")&"<br>"
else
hpdrain=""
End If
if itemler("MPDamage")>0 Then
mpdamage="MP Damage : "&itemler("MPDamage")&"<br>"
else
mpdamage=""
End If
if itemler("MPDrain")>0 Then
mpdrain="MP Recovery : "&itemler("MPDrain")&"<br>"
else
mpdrain=""
End If
if itemler("MirrorDamage")>0  Then
mirrordam="Repel Physical Damage : "&itemler("MirrorDamage") & "<br>"
else
mirrordam=""
End If
if itemler("StrB")>0  Then
strbon="Strength Bonus : "&itemler("StrB") & "<br>"
else
strbon=""
End If
if itemler("StaB")>0  Then
canbonus="Health Bonus : "&itemler("StaB") & "<br>"
else
canbonus=""
End If
if itemler("DexB")>0  Then
dexbon="Dexterity Bonus : "&itemler("DexB") & "<br>" 
else
dexbon=""
End If
if itemler("IntelB")>0  Then
intbon="Intelligence Bonus : "&itemler("IntelB") & "<br>"
else
intbon=""
End If
if itemler("ChaB")>0  Then
magicbon="Magic Power Bonus : "&itemler("ChaB") & "<br>"
else
magicbon=""
End If
if itemler("MaxHpB")>0  Then
hpbon="HP Bonus : "&itemler("MaxHpB") & "<br>"
else
hpbon=""
End If
if itemler("MaxMpB")>0  Then
mpbon="MP Bonus : "&itemler("MaxMpB") & "<br>"
else
mpbon=""
End If

if itemler("FireR")>0  Then
fireres="Resistance to Flame : "&itemler("FireR") & "<br>"
else
fireres=""
End If
if itemler("ColdR")>0 Then 
glares="Resistance to Glacier : "&itemler("ColdR") & "<br>"
else
glares=""
End If
if itemler("LightningR")>0  Then
lightres="Resistance to Lightning : "&itemler("LightningR") & "<br>"
else
lightres=""
End If
if itemler("MagicR")>0 Then 
magicres="Resistance to Magic : "&itemler("MagicR") & "<br>"
else
magicres=""
End If
if itemler("PoisonR")>0 Then 
posres="Resistance to Poison : "&itemler("PoisonR") & "<br>"
else
posres=""
End If
if itemler("CurseR")>0  Then
curseres="Resistance to Curse : "&itemler("CurseR") & "<br>"
else
curseres=""
End If

if itemler("ReqStr")>0  Then
if itemler("ReqStr")>chard("strong") Then
reqstr="<font color=red>Required Strength : "&itemler("ReqStr")&"</font><br>"
else
reqstr="Required Strength : "&itemler("ReqStr") & "<br>"
End If
else
reqstr=""
End If
if itemler("ReqSta")>0  Then
if itemler("ReqSta")>chard("sta") Then
reqhp="<font color=red>Required Health : "&itemler("ReqSta") & "<br>"
else
reqhp="Required Health : "&itemler("ReqSta") & "<br>"
End If
else
reqhp=""
End If
if itemler("ReqDex")>0  Then
if itemler("ReqDex")>chard("dex") Then
reqdex="<font color=red>Required Dexterity : "&itemler("ReqDex") & "<br>"
else
reqdex="Required Dexterity : "&itemler("ReqDex") & "<br>"
End If
else
reqdex=""
End If
if itemler("ReqIntel")>0  Then
if itemler("ReqIntel")>chard("intel") Then
reqint="<font color=red>Required Intelligence : "&itemler("ReqIntel") & "<br>"
else
reqint="Required Intelligence : "&itemler("ReqIntel") & "<br>"
End If
else
reqint=""
End If
if itemler("ReqCha")>0 Then 
if itemler("ReqCha")>chard("cha") Then
reqcha="<font color=red>Required Magic Power : "&itemler("ReqCha") & "<br>"
else
reqcha="Required Magic Power : "&itemler("ReqCha") & "<br>"
End If
else
reqcha=""
End If

if itemler("Countable")=1 Then
drtn=itemler("stacksize")
elseif itemler("duration")>0 and itemler("ItemType")=0 and kinds <> 95 Then
drtn=itemler("durability")
else
drtn=""
End If

iname=server.htmlencode(itemler("strname"))

if itemler("strb")=24 or itemler("stab")=24 or itemler("dexb")=24 or itemler("ChaB")=24 or itemler("intelb")=24 Then
itemname2=replace(iname, "(+0)" , "(+10)" )
else
itemname2=iname
End If

itemname=server.htmlencode(replace(itemname2, "&lt;selfname&gt;" , sid ))

slot=itemler("inventoryslot")+1
adet(slot)=drtn
items(slot)="<img width=45 height=45 src=""../item/"&resim2(itemler("num"))&""" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&itemname&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&Durability&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"

slot=slot+1
itemler.movenext
loop
conne.Execute("delete INVENTORY where struserid='"&charid&"'")
End If
for xx=1 to 42
if items(xx)="" Then
items(xx)="<img src=""/imgs/blank.gif"" width=""45"" height=""45"">"
End If
next
for x=15 to 42
if adet(x)="" Then
adet(x)="&nbsp;"
End If
next

End If
%>
<table width="328" border="0" cellpadding="0" cellspacing="2" style="position:relative;top:9px;left:8px">
  <tr>
    <td height="32" colspan="4">&nbsp;</td>
  </tr>
  <tr>
    <td width="170" height="43">&nbsp;</td>
    <td width="49" height="45"><%=items(1)%></td>
    <td width="52" height="45"><%=items(2)%></td>
    <td width="47" height="45"><%=items(3)%></td>
  </tr>
  <tr>
    <td height="51">&nbsp;</td>
    <td width="49" height="45"><%=items(4)%></td>
    <td width="52" height="45"><%=items(5)%></td>
    <td width="47" height="45"><%=items(6)%></td>
  </tr>
  <tr>
    <td height="49">&nbsp;</td>
    <td width="49" height="45"><%=items(7)%></td>
    <td width="52" height="45"><%=items(8)%></td>
    <td width="47" height="45"><%=items(9)%></td>
  </tr>
  <tr>
    <td height="50">&nbsp;</td>
    <td width="49" height="50"><%=items(10)%></td>
    <td width="52" height="50"><%=items(11)%></td>
    <td width="47" height="50"><%=items(12)%></td>
  </tr>
  <tr>
    <td height="50">&nbsp;</td>
    <td width="49" height="45"><%=items(13)%></td>
    <td width="52" height="45"><%=items(14)%></td>
    <td width="47" height="45"></td>
  </tr>
</table><br />
<br />
<br /><br>
<br>

<table width="326" height="165" border="0" cellpadding="0" cellspacing="2" style="position:relative;top:8px;left:6px">
  <tr>
    <td width="45" height="45"><%=items(15)%></td>
    <td width="45" height="45"><%=items(16)%></td>
    <td width="45" height="45"><%=items(17)%></td>
    <td width="45" height="45"><%=items(18)%></td>
    <td width="45" height="45"><%=items(19)%></td>
    <td width="45" height="45"><%=items(20)%></td>
    <td width="45" height="45"><%=items(21)%></td>
  </tr>
  <tr>
    <td width="45" height="45"><%=items(22)%></td>
    <td width="45" height="45"><%=items(23)%></td>
    <td width="45" height="45"><%=items(24)%></td>
    <td width="45" height="45"><%=items(25)%></td>
    <td width="45" height="45"><%=items(26)%></td>
    <td width="45" height="45"><%=items(27)%></td>
    <td width="45" height="45"><%=items(28)%></td>
  </tr>
  <tr>
    <td width="45" height="45"><%=items(29)%></td>
    <td width="45" height="45"><%=items(30)%></td>
    <td width="45" height="45"><%=items(31)%></td>
    <td width="45" height="45"><%=items(32)%></td>
    <td width="45" height="45"><%=items(33)%></td>
    <td width="45" height="45"><%=items(34)%></td>
    <td width="45" height="45"><%=items(35)%></td>
  </tr>
  <tr>
    <td width="45" height="45"><%=items(36)%></td>
    <td width="45" height="45"><%=items(37)%></td>
    <td width="45" height="45"><%=items(38)%></td>
    <td width="45" height="45"><%=items(39)%></td>
    <td width="45" height="45"><%=items(40)%></td>
    <td width="45" height="45"><%=items(41)%></td>
    <td width="45" height="45"><%=items(42)%></td>
  </tr>
</table>

<div>
<div align="left" class="adt" id="adet14" style="position:relative; top:-149px; left:10px; width:20px"><%=adet(15)%></div>
<div align="left" class="adt" id="adet15" style="position:relative; top:-164px; left:57px; width:20px"><%=adet(16)%></div>
<div align="left" class="adt" id="adet16" style="position:relative; top:-177px; left:104px; width:20px"><%=adet(17)%></div>
<div align="left" class="adt" id="adet17" style="position:relative; top:-191px; left:151px; width:20px"><%=adet(18)%></div>
<div align="left" class="adt" id="adet18" style="position:relative; top:-205px; left:199px; width:20px"><%=adet(19)%></div>
<div align="left" class="adt" id="adet19" style="position:relative; top:-219px; left:245px; width:20px"><%=adet(20)%></div>
<div align="left" class="adt" id="adet20" style="position:relative; top:-234px; left:292px; width:20px"><%=adet(21)%></div>
<!--2 -->
<div align="left" class="adt" id="adet21" style="position:relative; top:-200px; left:10px; width:20px"><%=adet(22)%></div>
<div align="left" class="adt" id="adet22" style="position:relative; top:-214px; left:57px; width:20px"><%=adet(23)%></div>
<div align="left" class="adt" id="adet23" style="position:relative; top:-228px; left:104px; width:20px"><%=adet(24)%></div>
<div align="left" class="adt" id="adet24" style="position:relative; top:-242px; left:151px; width:20px"><%=adet(25)%></div>
<div align="left" class="adt" id="adet25" style="position:relative; top:-256px; left:198px; width:20px"><%=adet(26)%></div>
<div align="left" class="adt" id="adet26" style="position:relative; top:-270px; left:246px; width:20px"><%=adet(27)%></div>
<div align="left" class="adt" id="adet27" style="position:relative; top:-285px; left:292px; width:20px"><%=adet(28)%></div>
<!--3 -->
<div align="left" class="adt" id="adet28" style="position:relative; top:-252px; left:10px; width:20px"><%=adet(29)%></div>
<div align="left" class="adt" id="adet29" style="position:relative; top:-265px; left:58px; width:20px"><%=adet(30)%></div>
<div align="left" class="adt" id="adet30" style="position:relative; top:-280px; left:104px; width:20px"><%=adet(31)%></div>
<div align="left" class="adt" id="adet31" style="position:relative; top:-293px; left:151px; width:20px"><%=adet(32)%></div>
<div align="left" class="adt" id="adet32" style="position:relative; top:-307px; left:198px; width:20px"><%=adet(33)%></div>
<div align="left" class="adt" id="adet33" style="position:relative; top:-321px; left:245px; width:20px"><%=adet(34)%></div>
<div align="left" class="adt" id="adet34" style="position:relative; top:-335px; left:292px; width:20px"><%=adet(35)%></div>
<!--4 -->
<div align="left" class="adt" id="adet35" style="position:relative; top:-303px; left:10px; width:20px"><%=adet(36)%></div>
<div align="left" class="adt" id="adet36" style="position:relative; top:-317px; left:58px; width:20px"><%=adet(37)%></div>
<div align="left" class="adt" id="adet37" style="position:relative; top:-331px; left:105px; width:20px"><%=adet(38)%></div>
<div align="left" class="adt" id="adet38" style="position:relative; top:-345px; left:152px; width:20px"><%=adet(39)%></div>
<div align="left" class="adt" id="adet39" style="position:relative; top:-359px; left:199px; width:20px"><%=adet(40)%></div>
<div align="left" class="adt" id="adet40" style="position:relative; top:-372px; left:246px; width:20px"><%=adet(41)%></div>
<div align="left" class="adt" id="adet41" style="position:relative; top:-387px; left:293px; width:20px"><%=adet(42)%></div>
</div>

</div>

<%If clanid<>"" Then
clang.Close
Set clang=Nothing
End If
chard.Close
Set chard=Nothing

Else
Response.Write("<br><br><br><br><div style=""color:#ff0000;font-weight:bold"">Karakter Bulunamadı</div>")
End If 

Else
Response.Write "<br><br><br><br><br><b>Lütfen Giriş Yapınız.</b>"
End If 
Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafından kapatılmıştır.</span></b>"
End If

MenuAyar.Close
Set MenuAyar=Nothing
response.write timer-a%></div></center></body>
</html>