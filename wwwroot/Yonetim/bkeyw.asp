<% if Session("durum")="esp" Then %>
<!--#include file="../function.asp"-->
<!--#include file="../_inc/conn.asp"-->

<%
itemname=request("keyw")
itemtype=request("itemtype")
derece=request("derece")
cls=request("class")
cset=request("set")
bonus=request("bonus")

if len(itemname)>=0 Then

if itemtype="hepsi" Then
itemtype=""
else
itemtype="and itemtype='"&ItemType&"'"
End If

if derece="hepsi" Then
derece=""
else
derece="and strname like '%+"&derece&"%'"
End If

if cls="hepsi" Then
clas=""
else
clas="and kind="&cls
End If

if cset="hepsi" Then
cset=""

elseif cset="shell" and cls="230" Then
csets="and strname like '%Complete%' and itemtype=5"
elseif cset="shell" Then
csets="and strname like '%shell%' and itemtype=5"

elseif cset="chitin" and cls="230" Then
csets="and strname like '%crimson%' and itemtype=5"
elseif cset="chitin" Then
csets="and strname like '%chitin%' and strname not like '%shell%' and itemtype=5"

elseif cset="fullplate" and cls="230" Then
csets="and strname like '%crystal%' and itemtype=5"
elseif cset="fullplate" Then
csets="and strname like '%ull plate%' and itemtype=5"
End If

if bonus="hepsi" Then
bonus=""
elseif bonus="str" Then
bonus="and StrB>0"
elseif bonus="dex" Then
bonus="and DexB>0"
elseif bonus="hp" Then
bonus="and StaB>0"
elseif bonus="mp" Then
bonus="and IntelB>0"
End If


set det=Conne.Execute("select * from item where strname like '%"&itemname&"%' "&itemtype&" "&derece&" "&clas&" "&csets&" "&bonus&" order by strname")
if not det.eof Then
do while not det.eof
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
renk="#CE8DC5"
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
Response.Write "<a href=""#"" onclick=""javascript:itemekle('"&resim2(det("num"))&"','"&det("num")&"','"&det("Duration")&"','1','"&trim(det("strname"))&"','"&("<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&server.htmlencode(det("strname"))&"<br>("&dtype&")</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>")&"',"&det("slot")&");return false"" id=""<img src=../item/"&resim(det("num"))&">"" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><img src=../item/"&resim2(det("num"))&"><br><font style=font-size:11px color="&renk&">"&server.htmlencode(det("strname"))&"<br>("&dtype&")</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);return false"" onMouseOut=""return nd();"">"&det("strname")&"&nbsp;&nbsp;"&det("num")&"<br></a>"

det.movenext
loop
End If

else
Response.Write "En az 2 Karakter yazın"
End If
%>
<script language="JavaScript" type="text/javascript">
function Sec(num)
{
document.search.keyw.value=num;

}
</script>
<% End If %>