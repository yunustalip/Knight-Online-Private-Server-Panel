<!--#include file="../_inc/conn.asp"-->
<!--#include file="../function.asp"-->
<% response.expires=0
if Session("durum")="esp" Then
Dim charid
charid=secur(Request.Querystring("charid"))%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Arama Sayfası</title>
<script type="text/javascript">
function itemozell(charid,inventoryslot){
  $.ajax({
	type: 'GET',
	url: 'invitem.asp?charid=<%=charid%>&inventoryslot='+inventoryslot,
	success: function(ajaxCevap) {
		$('div#itemoz').html(ajaxCevap);
	}
  });
}



</script>

</head>

<body>
<%
Response.Charset = "iso-8859-9"

dim items(42)
dim adet(42)
set userdt = Conne.Execute("SELECT * FROM USERDATA WHERE strUserId='"&charid&"'")
if userdt.Eof Then
Response.Write("<font color=""#FF0000""><strong>Karakter Bulunamadı</strong></font>")
Response.End
End If
conne.Execute("truncate table inventory_edit")
conne.Execute("Exec item_decode2 '"&charid&"'")
Dim itemler,userdt
set itemler=Conne.Execute("SELECT * FROM inventory_edit WHERE struserid='"&charid&"'")


%>  <table width="100%" border="0" align="right" bordercolor="333333" cellpadding="0" cellspacing="0" style="color:#FFFFFF; position:relative;">
            <tr>
			<td width="350" height="600" valign="top">
<div align="left" id="inven" style="color:#FFFFFF;font-size:11px;font-family:Arial,Verdana,Helvetica, sans-serif;font-weight:bold;background:url(../Item/inventory.gif);background-repeat:no-repeat; height:100%;width:350;">
              <% 
if not itemler.eof Then
Dim x,itemi,det,dtype,speed,kinds,delay,kind,renk,atack,weight,durability,duration,defans,dodging,incap,daggerac,swordac,clubac,axeac,spearac,bowac,firedam,icedam,ligthdam,posdam,hpdrain,mpdamage,mpdrain,mirrordam,strbon,canbonus,hpbon,dexbon,intbon,mpbon,magicbon,fireres,glares,lightres,magicres,posres,curseres,ReqStr,reqhp,reqdex,Reqint,Reqcha,iname,itemname2,drtn,sid,itemname,slot,maceac,xx
for x=0 to 41
itemi= x
set det=Conne.Execute("select * from ITEM where num='"&itemler("num")&"'")
if not det.eof Then

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
delay=""
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


if det("Damage")>0 Then 
atack="Attack Power : "&det("Damage") & "<br>"
else
atack=""
End If
if det("Weight")>0 Then 
weight="Weight : "&det("Weight") & "<br>"
else
weight=""
End If
if det("Duration")>1 and det("ItemType")=0 Then
duration="Quantity : "&det("Duration") & "<br>"
elseif det("Duration")>1  Then 
duration="Max Durability : "&det("Duration") & "<br>"
else
duration=""
End If
if det("Ac")>0 Then 
defans="Defense Ability : "&det("Ac") & "<br>"
else
defans=""
End If
if det("Evasionrate")>0 Then
dodging="Increase Dodging Power by : "&det("Evasionrate")&"<br>"
else
dodging=""
End If
if det("Hitrate")>0 Then
incap="Increase Attack Power by  : "&det("Hitrate")&"<br>"
else
incap=""
End If


if det("DaggerAc")>0 Then 
daggerac="Defense Ability (Dagger) : "&det("DaggerAc") & "<br>"
else
daggerac=""
End If
if det("SwordAc")>0 Then 
swordac="Defense Ability (Sword) : "&det("SwordAc") & "<br>"
else
swordac=""
End If
if det("MaceAc")>0 Then 
clubac="Defense Ability (Club) : "&det("MaceAc") & "<br>"
else
clubac=""
End If
if det("AxeAc")>0 Then 
axeac="Defense Ability (Axe) : "&det("AxeAc") & "<br>"
else
axeac=""
End If
if det("SpearAc")>0 Then 
spearac="Defense Ability (Spear) : "&det("SpearAc") & "<br>"
else
spearac=""
End If
if det("BowAc")>0 Then 
bowac="Defense Ability (Arrow) : "&det("BowAc") & "<br>"
else
bowac=""
End If
if det("FireDamage")>0 Then 
firedam="Flame Damage : "&det("FireDamage") & "<br>"
else
firedam=""
End If
if det("IceDamage")>0 Then 
icedam="Glacier Damage : "&det("IceDamage") & "<br>"
else
icedam=""
End If
if det("LightningDamage")>0 Then 
ligthdam="Lightning Damage : "&det("LightningDamage") & "<br>"
else
ligthdam=""
End If
if det("PoisonDamage")>0 Then 
posdam="Poison Damage : "&det("PoisonDamage") & "<br>"
else
posdam=""
End If
if det("HPDrain")>0 Then
hpdrain="HP Recovery : "&det("HPDrain")&"<br>"
else
hpdrain=""
End If
if det("HPDrain")>0 Then
mpdamage="MP Damage : "&det("MPDamage")&"<br>"
else
mpdamage=""
End If
if det("HPDrain")>0 Then
mpdrain="MP Recovery : "&det("MPDrain")&"<br>"
else
mpdrain=""
End If
if det("MirrorDamage")>0 Then 
mirrordam="Repel Physical Damage : "&det("MirrorDamage") & "<br>"
else
mirrordam=""
End If
if det("StrB")>0 Then 
strbon="Strength Bonus : "&det("StrB") & "<br>"
else
strbon=""
End If
if det("StaB")>0 Then 
canbonus="Health Bonus : "&det("StaB") & "<br>"
else
canbonus=""
End If
if det("MaxHpB")>0 Then 
hpbon="HP Bonus : "&det("MaxHpB") & "<br>"
else
hpbon=""
End If
if det("DexB")>0 Then 
dexbon="Dexterity Bonus : "&det("DexB") & "<br>" 
else
dexbon=""
End If
if det("IntelB")>0 Then 
intbon="Intelligence Bonus : "&det("IntelB") & "<br>"
else
intbon=""
End If
if det("MaxMpB")>0 Then 
mpbon="MP Bonus : "&det("MaxMpB") & "<br>"
else
mpbon=""
End If
if det("ChaB")>0 Then 
magicbon="Magic Power Bonus : "&det("ChaB") & "<br>"
else
magicbon=""
End If
if det("FireR")>0 Then 
fireres="Resistance to Flame : "&det("FireR") & "<br>"
else
fireres=""
End If
if det("ColdR")>0 Then 
glares="Resistance to Glacier : "&det("ColdR") & "<br>"
else
glares=""
End If
if det("LightningR")>0 Then 
lightres="Resistance to Lightning : "&det("LightningR") & "<br>"
else
lightres=""
End If
if det("MagicR")>0 Then 
magicres="Resistance to Magic : "&det("MagicR") & "<br>"
else
magicres=""
End If
if det("PoisonR")>0 Then 
posres="Resistance to Poison : "&det("PoisonR") & "<br>"
else
posres=""
End If
if det("CurseR")>0 Then 
curseres="Resistance to Curse : "&det("CurseR") & "<br>"
else
curseres=""
End If

if det("ReqStr")>0 Then 
if det("ReqStr")>userdt("strong") Then
reqstr="<font color=red>Required Strength : "&det("ReqStr")&"</font><br>"
else
reqstr="Required Strength : "&det("ReqStr") & "<br>"
End If
else
reqstr=""
End If
if det("ReqSta")>0 Then 
if det("ReqSta")>userdt("sta") Then
reqhp="<font color=red>Required Health : "&det("ReqSta") & "<br>"
else
reqhp="Required Health : "&det("ReqSta") & "<br>"
End If
else
reqhp=""
End If
if det("ReqDex")>0 Then 
if det("ReqDex")>userdt("dex") Then
reqdex="<font color=red>Required Dexterity : "&det("ReqDex") & "<br>"
else
reqdex="Required Dexterity : "&det("ReqDex") & "<br>"
End If
else
reqdex=""
End If
if det("ReqIntel")>0 Then 
if det("ReqIntel")>userdt("intel") Then
reqint="<font color=red>Required Intelligence : "&det("ReqIntel") & "<br>"
else
reqint="Required Intelligence : "&det("ReqIntel") & "<br>"
End If
else
reqint=""
End If
if det("ReqCha")>0 Then 
if det("ReqCha")>userdt("cha") Then
reqcha="<font color=red>Required Magic Power : "&det("ReqCha") & "<br>"
else
reqcha="Required Magic Power : "&det("ReqCha") & "<br>"
End If
else
reqcha=""
End If

iname=server.htmlencode(det("strname"))

if det("strb")=24 or det("stab")=24 or det("dexb")=24 or det("ChaB")=24 or det("intelb")=24 Then
itemname2=replace(iname, "(+0)" , "(+10)" )
else
itemname2=iname
End If

if det("Countable")=1 Then
drtn=itemler("stacksize")
elseif det("duration")>0 and det("ItemType")=0 and det("kind") <> 95 Then
drtn=itemler("durability")
else
drtn=""
End If

itemname=server.htmlencode(replace(itemname2, "&lt;selfname&gt;" , sid ))

slot=itemler("inventoryslot")+1
adet(slot)=drtn
items(slot)="<img id=i"&slot&" width=45 height=45 src=../item/"&resim2(det("num"))&" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&server.htmlencode(itemname)&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"

End If

itemler.movenext
next

for xx=1 to 42
if items(xx)="" Then
items(xx)="<img src=""../imgs/blank.gif"" width=""45"" height=""45"">"
End If
next
for x=15 to 42
if adet(x)="" Then
adet(x)="&nbsp;"
End If
next
%>
<table width="160" border="1" bordercolor="#333333" bordercolordark="#333333" bordercolorlight="#333333" cellpadding="2" cellspacing="1" style="position:relative;left:177px;top:38px">
  <tr>
    <td align="center" valign="middle" width="52" height="45" id="0" style="border:1px solid #00FF00" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','0'"%>)"><%=items(1)%></td>
    <td align="center" valign="middle"  width="48" height="45" id="1" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','1'"%>)"><%=items(2)%></td>
    <td align="center" valign="middle"  width="53" height="45" id="2" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','2'"%>)"><%=items(3)%></td>
  </tr>
  <tr>
    <td align="center" valign="middle"  width="52" height="51" id="3" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','3'"%>)"><%=items(4)%></td>
    <td align="center" valign="middle"  width="48" height="51" id="4" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','4'"%>)"><%=items(5)%></td>
    <td align="center" valign="middle"  width="53" height="51" id="5" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','5'"%>)"><%=items(6)%></td>
  </tr>
  <tr>
    <td align="center" valign="middle"  width="52" height="49" id="6" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','6'"%>)"><%=items(7)%></td>
    <td align="center" valign="middle"  width="48" height="49" id="7" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','7'"%>)"><%=items(8)%></td>
    <td align="center" valign="middle"  width="53" height="49" id="8" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','8'"%>)"><%=items(9)%></td>
  </tr>
  <tr>
    <td align="center" valign="middle"  width="52" height="50" id="9" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','9'"%>)"><%=items(10)%></td>
    <td align="center" valign="middle"  width="48" height="50" id="10" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','10'"%>)"><%=items(11)%></td>
    <td align="center" valign="middle"  width="53" height="50" id="11" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','11'"%>)"><%=items(12)%></td>
  </tr>
  <tr>
    <td align="center" valign="middle"  width="52" height="50" id="12" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','12'"%>)"><%=items(13)%></td>
    <td align="center" valign="middle"  width="48" height="50" id="13" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','13'"%>)"><%=items(14)%></td>
    <td width="53" height="50"></td>
  </tr>
</table>
<br />
<br />
&nbsp;&nbsp;<input type="button" value="KAYDET" disabled="disabled" onClick="dis();itemkayitall();" id="but" style="width:160px;height:35px">
<br>
<table width="330" height="166" border="1" bordercolor="#333333" cellpadding="0" cellspacing="0" style="position:relative;left:6px;top:33px;">
  <tr >
    <td width="45" height="45" id="14" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','14'"%>)"><%=items(15)%></td>
    <td width="45" height="45" id="15" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','15'"%>)"><%=items(16)%></td>
    <td width="45" height="45" id="16" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','16'"%>)"><%=items(17)%></td>
    <td width="45" height="45" id="17" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','17'"%>)"><%=items(18)%></td>
    <td width="45" height="45" id="18" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','18'"%>)"><%=items(19)%></td>
    <td width="45" height="45" id="19" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','19'"%>)"><%=items(20)%></td>
    <td width="45" height="45" id="20" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','20'"%>)"><%=items(21)%></td>
  </tr>
  <tr >
    <td  width="45" height="45" id="21" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','21'"%>)"><%=items(22)%></td>
    <td width="45" height="45" id="22" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','22'"%>)"><%=items(23)%></td>
    <td width="45" height="45" id="23" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','23'"%>)"><%=items(24)%></td>
    <td width="45" height="45" id="24" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','24'"%>)"><%=items(25)%></td>
    <td width="45" height="45" id="25" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','25'"%>)"><%=items(26)%></td>
    <td width="45" height="45" id="26" style="border:1px solid #4C4B36"  onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','26'"%>)"><%=items(27)%></td>
    <td width="45" height="45" id="27" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','27'"%>)"><%=items(28)%></td>
  </tr>
  <tr>
    <td width="45" height="45" id="28" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','28'"%>)"><%=items(29)%></td>
    <td width="45" height="45" id="29" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','29'"%>)"><%=items(30)%></td>
    <td width="45" height="45" id="30" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','30'"%>)"><%=items(31)%></td>
    <td width="45" height="45" id="31" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','31'"%>)"><%=items(32)%></td>
    <td width="45" height="45" id="32" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','32'"%>)"><%=items(33)%></td>
    <td width="45" height="45" id="33" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','33'"%>)"><%=items(34)%></td>
    <td width="45" height="45" id="34" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','34'"%>)"><%=items(35)%></td>
  </tr>
  <tr>
    <td width="45" height="45" id="35" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','35'"%>)"><%=items(36)%></td>
    <td width="45" height="45" id="36" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','36'"%>)"><%=items(37)%></td>
    <td width="45" height="45" id="37" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','37'"%>)"><%=items(38)%></td>
    <td width="45" height="45" id="38" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','38'"%>)"><%=items(39)%></td>
    <td width="45" height="45" id="39" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','39'"%>)"><%=items(40)%></td>
    <td width="45" height="45" id="40" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','40'"%>)"><%=items(41)%></td>
    <td width="45" height="45" id="41" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);itemozell('<%Response.Write charid&"','41'"%>)"><%=items(42)%></td>
  </tr>
</table><style>.adt{height:14px}</style><div class="adt" id="adet14" style="position:relative; top:-124px; left:10px; width:20px"><%=adet(15)%></div>
<div class="adt" id="adet15" style="position:relative; top:-139px; left:57px; width:20px"><%=adet(16)%></div>
<div class="adt" id="adet16" style="position:relative; top:-152px; left:104px; width:20px"><%=adet(17)%></div>
<div class="adt" id="adet17" style="position:relative; top:-166px; left:151px; width:20px"><%=adet(18)%></div>
<div class="adt" id="adet18" style="position:relative; top:-180px; left:200px; width:20px"><%=adet(19)%></div>
<div class="adt" id="adet19" style="position:relative; top:-194px; left:245px; width:20px"><%=adet(20)%></div>
<div class="adt" id="adet20" style="position:relative; top:-209px; left:292px; width:20px"><%=adet(21)%></div>
<!--2 -->
<div class="adt" id="adet21" style="position:relative; top:-175px; left:10px; width:20px"><%=adet(22)%></div>
<div class="adt" id="adet22" style="position:relative; top:-189px; left:57px; width:20px"><%=adet(23)%></div>
<div class="adt" id="adet23" style="position:relative; top:-203px; left:104px; width:20px"><%=adet(24)%></div>
<div class="adt" id="adet24" style="position:relative; top:-217px; left:151px; width:20px"><%=adet(25)%></div>
<div class="adt" id="adet25" style="position:relative; top:-231px; left:198px; width:20px"><%=adet(26)%></div>
<div class="adt" id="adet26" style="position:relative; top:-245px; left:245px; width:20px"><%=adet(27)%></div>
<div class="adt" id="adet27" style="position:relative; top:-260px; left:292px; width:20px"><%=adet(28)%></div>
<!--3 -->
<div class="adt" id="adet28" style="position:relative; top:-228px; left:10px; width:20px"><%=adet(29)%></div>
<div class="adt" id="adet29" style="position:relative; top:-241px; left:58px; width:20px"><%=adet(30)%></div>
<div class="adt" id="adet30" style="position:relative; top:-256px; left:104px; width:20px"><%=adet(31)%></div>
<div class="adt" id="adet31" style="position:relative; top:-269px; left:151px; width:20px"><%=adet(32)%></div>
<div class="adt" id="adet32" style="position:relative; top:-284px; left:198px; width:20px"><%=adet(33)%></div>
<div class="adt" id="adet33" style="position:relative; top:-297px; left:245px; width:20px"><%=adet(34)%></div>
<div class="adt" id="adet34" style="position:relative; top:-310px; left:292px; width:20px"><%=adet(35)%></div>
<!--4 -->
<div class="adt" id="adet35" style="position:relative; top:-278px; left:10px; width:20px"><%=adet(36)%></div>
<div class="adt" id="adet36" style="position:relative; top:-292px; left:58px; width:20px"><%=adet(37)%></div>
<div class="adt" id="adet37" style="position:relative; top:-306px; left:105px; width:20px"><%=adet(38)%></div>
<div class="adt" id="adet38" style="position:relative; top:-320px; left:152px; width:20px"><%=adet(39)%></div>
<div class="adt" id="adet39" style="position:relative; top:-334px; left:199px; width:20px"><%=adet(40)%></div>
<div class="adt" id="adet40" style="position:relative; top:-347px; left:246px; width:20px"><%=adet(41)%></div>
<div class="adt" id="adet41" style="position:relative; top:-362px; left:293px; width:20px"><%=adet(42)%></div>

<div style="position:relative; top:-580px; left:235px; width:100px; color: rgb(247, 231, 33);font-weight:bold;"><%=userdt("gold")%></div>
<%ElseIf charid="" Then
Response.End
Else
Response.Write "Kullanıcı Bulunamadı!"
End If%>
</div>
           </td>
<td align="center" width="300" valign="top" style="padding-top:30px" height="100" style="position:relative;left:20px">
<div id="itemoz"><%set itemozel=Conne.Execute("select * from inventory_edit where inventoryslot=0 and StrUserId='"&charid&"'")
if not itemozel.eof Then
set itemad=Conne.Execute("select strname from item where num="&itemozel("num")&"")
End If
if not itemad.eof and not itemozel.eof Then
itema=secur(itemad("strname"))
else
itema="&nbsp;"
End If
Response.Write "<form action=""javascript:itemkayit();document.getElementById('but').disabled=false;"" name=""itemk"" id=""itemk"">"&vbcrlf
Response.Write "<div id=""itemmname""><b>"&itema&"</b></div>"&vbcrlf
Response.Write "<input type=""hidden"" value="""&charid&""" name=""charid"" >"&vbcrlf
Response.Write "<br>Item num: <input type=text value="&itemozel("num")&" name=""num"" id=""num"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"&vbcrlf
Response.Write "<br>Item Serial: <input type=text value="""&itemozel("strserial")&""" name=""serial"" id=""serial"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"&vbcrlf
Response.Write "<br>Item Durability: <input type=text value="&itemozel("durability")&" name=""dur"" id=""dur"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"&vbcrlf
Response.Write "<br>Item Adet: <input type=text value="&itemozel("stacksize")&" name=""stacksize"" id=""stacksize"" onkeyup=""stacksizeupdate('0',this.value)"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"&vbcrlf
Response.Write "<br>Item Slot: <input type=text value=""0"" name=""inventoryslot"" id=""inventoryslot"" onblur=""itemkayit();document.getElementById('but').disabled=false;"">"&vbcrlf

Response.Write "<br><br><a href=""#"" onclick=""javascript:itemsil('"&inventoryslot&"')"">ITEMI SIL</a></form>"%>
</div>
<script type="text/javascript">
function chng()
{
  $.ajax({
	type: 'POST',
	url: 'bkeyw.asp',
	data: $('#search').serialize(),
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}


</script>
<input type="hidden"  name="citemmname" value="" id="citemmname">
<input type="hidden" value="" name="cnum" id="cnum">
<input type="hidden" value="" name="cserial" id="cserial">
<input type="hidden" name="cdur" id="cdur" value="">
<input type="hidden" name="cstacksize" id="cstacksize" value="">
<input type="hidden" name="cinventoryslot" id="cinventoryslot" value="">
<input type="hidden" name="cicon" id="cicon" value="">
<br>
<form action="javascript:chng();"  method="post" id="search" name="search">
	<table>
	<tr align="center">
	<td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td>
	</tr>
	<tr><td>
	<select name="itemtype">
	<option value="hepsi">Hepsi</option>
	<option value="0">Non Upgrade Item</option>
	<option value="1">Magic Item</option>
	<option value="2">Rare Item</option>
	<option value="3">Craft Item</option>
	<option value="4">Unique Item</option>
	<option value="5">Upgrade Item</option>
	<option value="6">Event Item</option>
	</select></td>
	<td>
	<select name="derece">
	<option value="hepsi">Hepsi</option>
	<option value="0">+0</option>
	<option value="1">+1</option>
	<option value="2">+2</option>
	<option value="3">+3</option>
	<option value="4">+4</option>
	<option value="5">+5</option>
	<option value="6">+6</option>
	<option value="7">+7</option>
	<option value="8">+8</option>
	<option value="9">+9</option>
	<option value="10">+10</option>
	</select></td>
	<td><select name="class">
	<option value="hepsi">Hepsi</option>
	<option value="210">Warrior</option>
	<option value="220">Rogue</option>
	<option value="240">Priest</option>
	<option value="230">Mage</option>
	</select>
	</td>
	<td>
	<select name="set">
	<option value="hepsi">Hepsi</option>
	<option value="shell">Shell Set</option>
	<option value="chitin">Chitin Set</option>
	<option value="fullplate">Full Plate Set</option>
	</select>
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select>
	</td>
	</tr>
	<tr><td colspan="5">
      <input type="text" style="width:280px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
	</td>
	</tr>
	<tr><td colspan="5">
      <div id="sonuc" style="width:280px; background-color:silver;">.. Item İsmini Yazın</div>
    </td></tr> 
</table>
</form>

</td>
</tr>

</table>
</body>

</html>
<%Else
Response.Write("Giriş Yapınız")
End If %>
