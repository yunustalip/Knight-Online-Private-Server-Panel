<% if Session("durum")="esp" Then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<META http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Arama Sayfası</title>
<script type="text/javascript" src="../js/jquery.js"></script>
<script type="text/javascript">
function chng(val)
{
  $.ajax({
	type: 'GET',
	url: 'keyw.asp',
	data: 'keyw='+val,
	success: function(ajaxCevap) {
		$('#sonuc').html(ajaxCevap);
	}
  });
}
</script>
</head>

<body onload="document.search.keyw.focus();">
<table>
<tr valign="top">
<td>
<form action="javascript:chng(document.search.keyw.value);"  method="post" id="search" name="search">
<input type="text" style="width:298px;"  name="keyw" id="keyw" />
<input type="submit" onclick="chng(document.search.keyw.value);" value="Item Ara">
<br>
<div id="sonuc" style="width:304px; background-color:silver;">.. yazmaya başlayın</div>
</form>
</td>
<td>

<!--#include file="../function.asp"-->
 <%
Response.Charset = "iso-8859-9"
charid=secur(Request.Querystring("sid"))
set rs1 = conne.Execute("delete INVENTORY_EDIT where struserid='"&charid&"'")
Set rs = conne.Execute("Exec item_decode '"&charid&"'")

set itemler=Conne.Execute("SELECT * FROM INVENTORY_EDIT WHERE struserid='"&charid&"'")
set money = Conne.Execute("SELECT * FROM USERDATA WHERE strUserId='"&charid&"'")

%><form action="itemky.asp?charid=<%=charid%>" method="post">
 <table border="1" align="right" style=" background:#333333;color:#FFFFFF" >

  <tr>
<% 
if not itemler.eof Then
do while itemler("Inventoryslot")<42

Set item = Conne.Execute("SELECT * FROM item WHERE num='"&itemler("num")&"'")
set det=Conne.Execute("select * from ITEM where num='"&itemler("num")&"'")

if not item.eof Then

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
dtype="Non Upgrade Item"
renk="white"
elseif dtype=1 Then
dtype="Magic Item"
renk="blue"
elseif dtype=2 Then
dtype="Rare Item"
renk="yellow"
elseif dtype=3 Then
dtype="Craft Item"
renk="lime"
elseif dtype=4 Then
dtype="Unique Item"
renk="#DFC68C"
elseif dtype=5 Then
dtype="Upgrade Item"
renk="purple"
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
reqstr="Required Strength : "&det("ReqStr") & "<br>"
else
reqstr=""
End If
if det("ReqSta")>0 Then 
reqhp="Required Health : "&det("ReqSta") & "<br>"
else
reqhp=""
End If
if det("ReqDex")>0 Then 
reqdex="Required Dexterity : "&det("ReqDex") & "<br>"
else
reqdex=""
End If
if det("ReqIntel")>0 Then 
reqint="Required Intelligence : "&det("ReqIntel") & "<br>"
else
reqint=""
End If
if det("ReqCha")>0 Then 
reqcha="Required Magic Power : "&det("ReqCha") & "<br>"
else
reqcha=""
End If


items="<img width=45 height=45 src=../item/"&resim2(det("num"))&" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&server.htmlencode(det("strname"))&"<br>("&dtype&")</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"&"<input type=text size=9 style='font-size:11px' name='"&itemler("inventoryslot")&"' value="&det("num")&">"



if itemler("InventorySlot")="3" or itemler("InventorySlot")="6" or itemler("InventorySlot")="9" or itemler("InventorySlot")="12" Then
Response.Write "</tr><tr><td><font size=2 face=times new roman>"&items&"</font></td>"
elseif itemler("InventorySlot")="14" Then
Response.Write "<td></td></tr><tr height=40><td colspan='7' align='right' valign='middle' style='padding-right:45px'><font size='1' color='#DFC68C' face='verdana'><b></b></font></td></tr><tr><td><font size=2 face=times new roman>"&items&"</font></td>"
elseif itemler("InventorySlot")="21" or itemler("InventorySlot")="28" or itemler("InventorySlot")="35" Then
Response.Write "</tr><tr><td><font size=2 face=times new roman>"&items&"</font></td>"
else
Response.Write "<td width=40><font size=2 face=times new roman>"&items&"</font></td>"
End If

elseif itemler("InventorySlot")="3" or itemler("InventorySlot")="6" or itemler("InventorySlot")="9" or itemler("InventorySlot")="12" Then
Response.Write "</tr><tr><td height=45 width=45><input type=text size=9 style='font-size:11px' name='"&itemler("inventoryslot")&"'></td>"
elseif itemler("InventorySlot")="14" Then
Response.Write "<td></td></tr><tr><td colspan='7' align='right'><font size='2'><b><center>Inventory</center></b></font></td></tr><tr><td><input type=text size=9 style='font-size:11px' name='"&itemler("inventoryslot")&"'></td>"
elseif itemler("InventorySlot")="21" or itemler("InventorySlot")="28" or itemler("InventorySlot")="35" Then
Response.Write "</tr><tr><td height=45 width=45><input type=text size=9 style='font-size:11px' name='"&itemler("inventoryslot")&"'></td>"
else Response.Write "<td height=45 width=45><input type=text size=9 style='font-size:11px' name='"&itemler("inventoryslot")&"'></td>"
End If 

itemler.movenext
loop

elseif charid="" Then
Response.End
else
Response.Write "Kullanıcı Bulunamadı!"
End If%>
</tr>
<tr>
<td><input type=hidden value="<%=charid%>"><input type="submit" value="KAYDET" /></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>

</html>
<% End If %>
