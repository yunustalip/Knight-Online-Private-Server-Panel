<% response.expires=0
if Session("durum")="esp" Then
Response.Charset = "iso-8859-9"
sayfa=secur(Request.Querystring("sayfa"))
if sayfa="" Then
sayfa=1
elseif sayfa<=0 or sayfa>8 or isnumeric(sayfa)=false Then
sayfa=1
else
sayfa=Request.Querystring("sayfa")
End If

sl=(sayfa)*24
sm=sayfa*24-24
%>
<html>

<head>
<META http-equiv="content-type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Arama Sayfası</title>
<script type="text/javascript" src="../js/jquery.js"></script>
<script type="text/javascript">

function itemozell(inventoryslot){
  $.ajax({
	type: 'GET',
	url: 'bankabaslangicitemleri.asp?inventoryslot='+inventoryslot,
	success: function(ajaxCevap) {
		$('div#itemoz').html(ajaxCevap);
	}
  });
}
 
    function loadpage(syf){
    $.ajax({
       type: 'GET',
       url: syf,
       start: $('#itemler').html('<center><img src="../imgs/38-1.gif"><br><b>Itemler Yükleniyor...</b></center>'),
       success: function(ajaxCevap) {
          $('#itemler').html(ajaxCevap);
       }
    });
    }
 function selectedbox(slot){
 for (var i=<%=sm%>; i<<%=sl%>; i++){
  var spn = document.getElementById(i);
  spn.style.border = "";
	 }
	document.getElementById(slot).style.border = "1px solid #00FF00";
document.getElementById('but').disabled=false
	}
    function itemekle(rsm,num,dur){
		var ids=document.getElementById('inventoryslot').value
		document.getElementById(ids).innerHTML='<img src="../item/'+rsm+'>';
		document.getElementById('num').value=num;
		document.getElementById('serial').value='0';
		document.getElementById('dur').value=dur;
		document.getElementById('stacksize').value='0';
		eval(itemkayt())
		}
		function itemsil(slot)
		{
		var ids=document.getElementById('inventoryslot').value
		document.getElementById(ids).innerHTML='&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
		document.getElementById('num').value=0;
		document.getElementById('serial').value='0';
		document.getElementById('dur').value=0;
		document.getElementById('stacksize').value='0';
		eval(itemkayt())
		}
	function itemkayt()
	{
	  $.ajax({
		type: 'post',
		url: 'bankabaslangicitemlerikaydet.asp?kyt=one',
		data: $('#itemk').serialize(),
	  });
	document.getElementById('but').disabled=false
	}
	function itemkayt2()
	{
	  $.ajax({
		type: 'post',
		url: 'bankabaslangicitemlerikaydet.asp?kyt=all',
		data: $('#itemk').serialize(),
	
	  });
	document.getElementById('but').disabled=true
	}
    </script>
</head>

<body onLoad="document.search.keyw.focus();" oncontextmenu="return false" >

<table name="itemlerim" id="">
  <tr valign="top" >
    <td rowspan="2"><!--#include file="../_inc/conn.asp"-->
        <!--#include file="../function.asp"-->
        <%

if Request.Querystring("sayfa")=""  Then
Conne.Execute("exec banka_baslangicitem_decode")
else
End If
set itemler=Conne.Execute("SELECT * FROM BANKA_CHECK WHERE straccountid='baslangic-item' order by inventoryslot")
if not itemler.eof Then
itemler.move((sayfa-1)*24)

%>
<table border="2" width="300" bgcolor="black">
<tr>
<%
for slot=1 to 24
itemi= (sayfa-1)*24+(slot-1)
set det=Conne.Execute("select itemtype,delay,num,kind,strname,damage,weight,duration,ac,Evasionrate,Hitrate,daggerac,swordac,axeac,spearac,bowac,maceac,firedamage,icedamage,LightningDamage,poisondamage,hpdrain,MPDamage,mpdrain,MirrorDamage,StrB,StaB,MaxHpB,DexB,IntelB,MaxMpB,ChaB,FireR,coldr,LightningR,magicr,poisonr,curser,reqstr,ReqSta,reqdex,reqintel,reqcha,Countable from ITEM where num='"&itemler("dwid")&"'")
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
elseif det("kind") = 98 Then
kind="Upgrade Scroll"
elseif det("kind") = 110 Then
kind="Staff"
elseif det("kind") = 120 Then
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

iname=server.htmlencode(det("strname"))
itemname=replace(iname, "(+0)" , "(+10)" )

if det("Damage")>0 Then 
atack="Attack Power : "&det("Damage") & "<br>"
else
atack=""
End If
if det("Weight")>1 Then 
weight="Weight : "&det("Weight") & "<br>"
else
weight=""
End If
if itemler("durability")>1 and not det("kind")=255 Then
Durability="Durability : "&itemler("durability")&"<br>"
else
Durability=""
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
if det("MPDamage")>0 Then
mpdamage="MP Damage : "&det("MPDamage")&"<br>"
else
mpdamage=""
End If
if det("MPDrain")>0 Then
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

if det("Countable")=1 Then
drtn=itemler("stacksize")
elseif det("duration")>0 and det("itemtype")=0 Then
drtn=itemler("durability")
else
drtn=""
End If

items="<img width=45 height=45 src=../item/"&resim2(det("num"))&" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&itemname&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&Durability&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', LEFT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"

Response.Write "<td width=""45"" height=""45"" id="""&itemi&"""  onclick=""selectedbox(this.id);itemozell('"&itemler("inventoryslot")&"')""> "&items&"</td>"
if slot mod 6=0 Then
Response.Write "</tr><tr>"
End If

else
Response.Write "<td width=""45"" height=""45"" id="""&itemi&""" onclick=""selectedbox(this.id);itemozell('"&itemler("inventoryslot")&"')"">&nbsp;</td>"
if slot mod 6=0 Then
Response.Write "</tr><tr>"
End If
End If

itemler.movenext
next
Response.Write "<tr><td colspan=""6"" align=""center"">"
if sayfa>1 Then 
Response.Write "<a href=""javascript:loadpage('bankabaslangicitemleribul.asp?sayfa="&sayfa-1&"')""><img src=""../imgs/isolok.gif"" border=""0""  align='absmiddle'></a>"
End If
for x=1 to 8
if x=cint(sayfa) Then
Response.Write "<font color=""#ffee2f"" face=""verdana"" size=""2""><b> "&x&" </b></font>"
else
Response.Write "<a href=""javascript:loadpage('bankabaslangicitemleribul.asp?sayfa="&x&"')""><font color=""#c3cccc"" face=""verdana"" size=""2""><b> "&x&" </b></font></a>"
End If

next
if sayfa<8 Then 
Response.Write "<a href=""javascript:loadpage('bankabaslangicitemleribul.asp?sayfa="&sayfa+1&"')""><img src=""../imgs/isagok.gif"" border=""0""  align='absmiddle'></a>"
End If
Response.Write "</td></tr>"
else
conne.Execute("delete banka_check where straccountid='baslangic-item'")
Conne.Execute("exec banka_baslangicitem_decode")
End If%>
            </tr>
			<tr><td colspan="6" align="center"><input type="button" onclick="itemkayt2();" id="but" value="KAYDET" style="width:275"></td></tr>
</table>
</td>
<td align="center" width="300" height="100"><div id="itemoz" ></div></td>
</tr>
<tr>
<td><br><script type="text/javascript">
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
itemozell('<%=itemi-23%>');
selectedbox('<%=itemi-23%>');

</script>
<form action="javascript:chng();"  method="post" id="search" name="search">
<table>
<tr valign="top">
    <td>
	<table >
	<tr align="center"><td>Item Türü</td><td>Grade</td><td>Class</td><td>Item Seti</td><td>Bonus</td></tr>
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
	</td>
	<td>
	<select name="bonus">
	<option value="hepsi">Hepsi</option>
	<option value="str">Str</option>
	<option value="dex">Dex</option>
	<option value="hp">Hp</option>
	<option value="mp">Mp</option>
	</select>
	</tr></table>
      <input type="text" style="width:200px;"  name="keyw" id="keyw" />
      <input name="submit" type="submit" value="Item Ara">
	</td>
	</tr>
	<tr><td>
      <div id="sonuc" style="width:280px; background-color:silver;">.. Item İsmini Yazın</div>
    </td>
	</tr> 
</table>
</form></td>
</tr>
</table>


</body>

</html>
<%End If %>