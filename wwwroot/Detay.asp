<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<!--#include file="Guvenlik.asp"-->
<%Dim MenuAyar,ksira
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then
if Session("durum")="game" and Session("accountid")<>"" Then
accid=Session("accountid")

elseif Session("login")="ok" and Session("username")<>"" Then
Dim Accid
Accid=Session("username")

Else
Server.Execute("hata.asp")
Response.End
End If
dim ItemId,pitem,det,cashp,dtype,speed,kinds,delay,kind,renk,atack,weight,durability,duration,defans,dodging,incap,daggerac,swordac,clubac,axeac,spearac,bowac,firedam,icedam,ligthdam,posdam,hpdrain,mpdamage,mpdrain,mirrordam,strbon,canbonus,hpbon,dexbon,intbon,mpbon,magicbon,fireres,glares,lightres,magicres,posres,curseres,ReqStr,reqhp,reqdex,Reqint,Reqcha,maceac,healtbon
itemid=secur(Request.Querystring("itemid"))
If Not len(itemid)>0 Then
Response.End
End If
If IsNumeric(itemid)=False Then
Response.Write "Unknown Item No"
Response.End
End If
Set pitem=Conne.Execute("select * from pus_itemleri where itemkodu='"&itemid&"'")

If pitem.eof Then
Response.Write "<br><br><font color=""white""><center>Item Bulunamadý!</center></font>"
Response.End
End If

Set det=Conne.Execute("select * from item where num='"&itemid&"'")
If det.Eof Then
Response.Write "<br><br><font color=""white""><center>Item Bulunamadý!</center></font>"
Response.End
End If
Set cashp=Conne.Execute("select * from tb_user where straccountid='"&accid&"'")
If cashp.Eof Then
Response.End
End If
dtype=det("ItemType")
speed=det("delay")

if speed>0 and speed<90 Then
delay = "Atack Speed : Very Fast<br>"
elseif speed>89 and speed<111 and not det("kind")=>91 and not det("kind")=<95 Then
delay = "Atack Speed : Fast<br>"
elseif speed>110 and speed<131 Then
delay = "Atack Speed : Normal<br>"
elseif speed>130 and speed<151 Then
delay = "Atack Speed : Slow<br>"
elseif speed>150 and speed<201 Then
delay = "Atack Speed : Very Slow<br>"
else
End If

if det("kind")=11 Then
kind="Dagger"
elseif det("kind") =21 Then
kind="One-handed Sword"
elseif det("kind") = 22 Then
kind="Two-handed Sword"
elseif det("kind") =31 Then
kind= "Axe"
elseif det("kind") = 32 Then
kind="Two-handed Axe"
elseif det("kind") = 41 Then
kind="Club"
elseif det("kind") = 42 Then
kind="Two-handed Club"
elseif det("kind") = 51 Then
kind="Spear"
elseif det("kind") = 52 Then
kind="Long Spear"
elseif det("kind") = 60 Then
kind="Shield"
elseif det("kind") = 70 Then
kind="Bow"
elseif det("kind") = 71 Then
kind="Crossbow"
elseif det("kind") = 91 Then
kind="Earring"
elseif det("kind") = 92 Then
kind="Necklace"
elseif det("kind") = 93 Then
kind="Ring"
elseif det("kind") = 94 Then
kind="Belt"
elseif det("kind") = 95 Then
kind="Lune Item"
elseif det("kind") = 110 Then
kind="Staff"
elseif det("kind") = 210 Then
kind="Warrior Armor"
elseif det("kind") = 220 Then
kind="Rogue Armor"
elseif det("kind") = 230 Then
kind="Magician Armor"
elseif det("kind") = 240 Then
kind="Priest Armor"
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


if det("Damage")>0 Then 
atack="Attack Power : "&det("Damage") & "<br>"
End If
if det("Weight")>0 Then 
weight="Weight : "&det("Weight") & "<br>"
End If
if det("Duration")>1 and det("ItemType")=0 Then
duration="Quantity : "&det("Duration") & "<br>"
elseif det("Duration")>1  Then 
duration="Max Durability : "&det("Duration") & "<br>"
End If
if det("Ac")>0 Then 
defans="Defense Ability : "&det("Ac") & "<br>"
End If
if det("Evasionrate")>0 Then
dodging="Increase Dodging Power by : "&det("Evasionrate")&"<br>"
End If
if det("Hitrate")>0 Then
incap="Increase Attack Power by  : "&det("Hitrate")&"<br>"
End If


if det("DaggerAc")>0 Then 
daggerac="Defense Ability (Dagger) : "&det("DaggerAc") & "<br>"
End If
if det("SwordAc")>0 Then 
swordac="Defense Ability (Sword) : "&det("SwordAc") & "<br>"
End If
if det("MaceAc")>0 Then 
clubac="Defense Ability (Club) : "&det("MaceAc") & "<br>"
End If
if det("AxeAc")>0 Then 
axeac="Defense Ability (Axe) : "&det("AxeAc") & "<br>" 
End If
if det("SpearAc")>0 Then 
spearac="Defense Ability (Spear) : "&det("SpearAc") & "<br>"
End If
if det("BowAc")>0 Then 
bowac="Defense Ability (Arrow) : "&det("BowAc") & "<br>"
End If
if det("FireDamage")>0 Then 
firedam="Flame Damage : "&det("FireDamage") & "<br>"
End If
if det("IceDamage")>0 Then 
icedam="Ice Damage : "&det("IceDamage") & "<br>"
End If
if det("LightningDamage")>0 Then 
ligthdam="Lightning Damage : "&det("LightningDamage") & "<br>"
End If
if det("PoisonDamage")>0 Then 
posdam="Poison Damage : "&det("PoisonDamage") & "<br>"
End If
if det("HPDrain")>0 Then
hpdrain="HP Recovery : "&det("HPDrain")&"<br>"
End If
if det("HPDrain")>0 Then
mpdamage="MP Damage : "&det("MPDamage")&"<br>"
End If
if det("HPDrain")>0 Then
mpdrain="MP Recovery : "&det("MPDrain")&"<br>"
End If
if det("MirrorDamage")>0 Then 
mirrordam="Repel Physical Damage : "&det("MirrorDamage") & "<br>"
End If
if det("StrB")>0 Then 
strbon="Strength Bonus : "&det("StrB") & "<br>"
End If
if det("StaB")>0 Then 
healtbon="Health Bonus : "&det("StaB") & "<br>"
End If
if det("MaxHpB")>0 Then 
hpbon="HP Bonus : "&det("MaxHpB") & "<br>"
End If
if det("DexB")>0 Then 
dexbon="Dexterity Bonus : "&det("DexB") & "<br>" 
End If
if det("IntelB")>0 Then 
intbon="Intelligence Bonus : "&det("IntelB") & "<br>"
End If
if det("MaxMpB")>0 Then 
mpbon="MP Bonus : "&det("MaxMpB") & "<br>"
End If
if det("ChaB")>0 Then 
magicbon="Magic Power Bonus : "&det("ChaB") & "<br>"
End If
if det("FireR")>0 Then 
fireres="Resistance to Flame : "&det("FireR") & "<br>"
End If
if det("ColdR")>0 Then 
glares="Resistance to Glacier : "&det("ColdR") & "<br>"
End If
if det("LightningR")>0 Then 
lightres="Resistance to Lightning : "&det("LightningR") & "<br>"
End If
if det("MagicR")>0 Then 
magicres="Resistance to Magic : "&det("MagicR") & "<br>"
End If
if det("PoisonR")>0 Then 
posres="Resistance to Poison : "&det("PoisonR") & "<br>"
End If
if det("CurseR")>0 Then 
curseres="Resistance to Curse : "&det("CurseR") & "<br>"
End If

if det("ReqStr")>0 Then 
reqstr="Required Strength : "&det("ReqStr") & "<br>"
End If
if det("ReqSta")>0 Then 
reqhp="Required Health : "&det("ReqSta") & "<br>"
End If
if det("ReqDex")>0 Then 
reqdex="Required Dexterity : "&det("ReqDex") & "<br>"
End If
if det("ReqIntel")>0 Then 
reqint="Required Intelligence : "&det("ReqIntel") & "<br>"
End If
if det("ReqCha")>0 Then 
reqcha="Required Magic Power : "&det("ReqCha") & "<br>"
End If%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head>

<title>Detay</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-2">
<link href="images/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="js/overlib.js"></script>
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

function buy_go(num) {
	location.href="buy.asp?itemid="+num;
}
function detail_go(num) {
	right.location.href="detail.asp?itemid="+num;
}

function buy_go_premium(itemid){
	document.frmbuyitem.itemid.value=itemid;
	document.frmbuyitem.action="/cash/premium_item_buy.asp";
	document.frmbuyitem.submit();
}

function wish_go(num) {
	location.href="cart.asp?itemid="+num;
}

function buy_gold(itemid){
	document.frmbuyitem.itemid.value=itemid;
	document.frmbuyitem.action="/cash/buy_gold.asp";
	document.frmbuyitem.submit();
}

function buy_go_transfer(itemid){
	document.frmbuyitem.itemid.value=itemid;
	document.frmbuyitem.action="/cash/transfer_item_buy.asp";
	document.frmbuyitem.submit();
}

function buy_go_genderchange(itemid){
	document.frmbuyitem.itemid.value=itemid;
	document.frmbuyitem.action="/cash/gender_change_buy.asp";
	document.frmbuyitem.submit();	
}


//-->
</script>
<style type="text/css">
body {background-color: transparent}
</style>

</head><body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="MM_preloadImages('/images/butt/butt_buy_on.gif','/images/butt/butt_cart_on.gif','/images/butt/butt_back_on.gif')">

<table border="0" cellpadding="0" cellspacing="0" width="288">
  <tbody><tr>
    <td height="31">
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

<table style="position: relative; top: -5px;" border="0" cellpadding="0" cellspacing="0">
  <tbody><tr>
        <td width="8"></td>
		<td style="padding-top: 0px;">
			<a href="cart.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_cart','','images/right_sub_memu_02.gif',1)">
			<img src="images/right_sub_memu_02.gif" name="butt_cart" border="0" width="100" height="29">			</a>		</td>
        <td width="1"></td>
	    <td style="padding-top: 0px;"><!--<a href="/imfor/detail.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_detail','','/include/images/right_sub_memu2_03.gif',1)">-->
		<img src="images/right_sub_memu_03b.gif" name="butt_detail" border="0" width="59" height="29">		</td>
		<td width="1"></td>
		<td style="padding-top: 0px;"><a href="http://k2shop.knightonlineworld.com/buy/promo.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_promo','','images/butt/butt_promo.gif',1)"><img style="display:none" src="images/butt_promo.gif" name="butt_promo" border="0" width="70" height="29"></a></td>
		
    </tr>
</tbody></table>

    </td>
  </tr>
  <tr>
    <td align="center" valign="top" width="288" height="473">
      <table border="0" cellpadding="0" cellspacing="0" width="230">
        <tbody><tr>
          <td height="28">&nbsp;</td>
        </tr>
        <tr>
          <td align="center">
		  <div id="item_list" style="border: 0px none ; overflow: auto; height: 425px;">
              <table border="0" cellpadding="0" cellspacing="0" width="220">
                <tbody><tr>
                  <td width="83"><table border="0" cellpadding="0" cellspacing="0" width="72" height="72">
                      <tbody><tr>
<td  onMouseOver="return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color=<%=renk%>><%=server.htmlencode(det("strname"))%><br><%=dtype%></font><br><font color=white style=font-size:11px><%=kind%></font><br><br></center><font color=white style=font-size:11px;><%=atack&delay&weight&duration&defans&dodging&incap%></font><font color=lime style=font-size:11px><%=daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&hpbon%><%=healtbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres%></font><font color=white style=font-size:11px><%=ReqStr&reqhp&reqdex&Reqint&Reqcha%></font>', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);" onMouseOut="return nd();"  align="center" background="images/table1_01.gif"><img src="item/<%=resim(pitem("resim"))%>" width="64" height="64"></td>
                      </tr>
                    </tbody></table></td>
                  <td><table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tbody><tr>
                        <td style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" height="28">*<strong>
                          <%=pitem("itemismi")%></strong></td>
                      </tr>
                      <tr>
                        <td style="font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;">* Ücret: <%=pitem("ucret")%>
                        </td>
                      </tr>
                      <tr>
                        <td style="font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;">* Adet: <%=pitem("adet")%></td>
                      </tr>
<%if pitem("premiumitem")=1 Then%>
                      <tr>
                        <td style="font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;">* Kullaným Günü: <%=pitem("kullanimgunu")%> Gün</td>
                      </tr>
<%End If%>
                    </tbody></table></td>
                </tr>
              </tbody></table>

            <table border="0" cellpadding="0" cellspacing="0" width="218">
              <tbody><tr>
                <td valign="top" height="10"></td>
                <td></td>
              </tr>
              <tr>
                <td valign="top" width="8"><img src="images/icon_02.gif" width="5" height="9"></td>
                <td><span style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;">Açýklama :</span>
                  <span style="font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;"><%=pitem("detay")%></span></td>
              </tr>
            </tbody></table>
            <br>
            <table border="0" cellpadding="0" cellspacing="0" width="220" height="35">
              <tbody><tr align="center" valign="bottom">
                <td>
	               
		                  <a href="javascript:buy_go(<%=itemid%>);" target="right" onMouseOver="MM_swapImage('butt_buy111','','images/butt/butt_buy_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt/butt_buy.gif" name="butt_buy111" id="butt_buy1" border="0" width="92" height="30"></a>                </td>
                <td>
                	<a href="javascript:wish_go(<%=itemid%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_cart1111','','images/butt/butt_cart_on.gif',1)"><img src="images/butt/butt_cart.gif" name="butt_cart1111" id="butt_cart1" border="0" width="92" height="30"></a><!-- ÂòÇÑ¸ñ·Ï º¸±â¿¡¼­ µé¾î°¬À»°æ¿ì <a href="/buy/cart.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_back','','/images/butt/butt_back_on.gif',1)"><img src="/images/butt/butt_back.gif" name="butt_back" width="92" height="30" border="0"></a>-->               	</td>
              </tr>
            </tbody></table></div></td>
        </tr>
      </tbody></table></td>
  </tr>
  <tr>
      <td style="padding-top: 10px;" align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="252" height="26"><tbody><tr><td style="background-image: url(images/n_table2_02.gif); background-repeat: no-repeat;" height="26">
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr>
  <td style="font-size: 11px; color: rgb(255, 238, 47); font-weight: bold;" align="right"><strong><%=cashp("cashpoint")%> P</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr></tbody></table>
</td></tr></tbody></table>
    </td>
  </tr>
  <tr>
    <td align="center" height="44">
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

<table border="0" cellpadding="0" cellspacing="0" width="254">
  <tbody><tr> 
    <td align="center"><a href="my_item.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_item','','images/butt/butt_item_on.gif',1)"><img src="images/butt/butt_item.gif" name="butt_item" border="0" width="119" height="30"></a></td>
    
  </tr>
</tbody></table>

    </td>
  </tr>
</tbody></table>

</body></html>
<%else
Response.Write "<br><b>Bu bölüm Admin tarafýndan kapatýlmýþtýr.</b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>