<!--#include file="../_inc/conn.asp"-->
<!--#include file="../Function.asp"-->
<!--#include file="../Guvenlik.asp"-->
<br /><img src="../imgs/pazar.gif" /><br /><br /><%ips=Request.ServerVariables("REMOTE_HOST")
Response.Charset = "iso-8859-9"
Response.expires=0 
Response.Write "<base href=""http://"&Request.ServerVariables("server_name")&""">"
Dim MenuAyar,ksira,ips,islem,chars
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='SellingPanel'")
If MenuAyar("PSt")=1 Then

If Session("login")="ok" Then

Action = Secur(Request.Querystring("Action"))

Select Case Action
Case ""

Set MyMoney=Conne.Execute("Select Sum(Price) as ToplamPara From SellingItems Where strAccountid='"&Session("username")&"' And Status=1	")
ToplamPara=MyMoney("ToplamPara")
If Isnull(ToplamPara) Then ToplamPara=0
%>
<br><br><table width="500">
	<tr>
		<td valign="top"  align="center"><a href="#" onclick="pageload('Sayfalar/SellingPanel.Asp?Action=Buy','1');return false" class="link1">
		<img src="imgs/itembuy.gif" border="0" align="top"></a>
		</td>
		<td valign="top"  align="center">
		<a href="#" onclick="pageload('Sayfalar/SellingPanel.Asp?Action=Sell','1');return false" class="link1">
		<img src="imgs/itemsell.gif" border="0" align="top"></a>
		<td>
	</tr>
	<tr>
	<td colspan="2">
	<table width="500" bordercolor="#cccc00" border="1" cellpadding="4" cellspacing="1" align="center" style="border-collapse: collapse;">
	<tr>
	<td align="center" colspan="3"><strong>Satýþ Bilgileri</strong></td>
	</tr>
		<tr>
	<td align="center"><strong>Kazanýlan Para: <%=ToplamPara%> </strong></td>
	<td align="center"><strong><a href="#" onclick="pageload('Sayfalar/SellingPanel.Asp?Action=TransportBank','1');return false" class="link1">Parayý Bankaya Aktar>></a></strong></td>
	</tr>
		<tr>
	<td align="center"><strong><a href="#" onclick="pageload('Sayfalar/SellingPanel.Asp?Action=MyItems&Durum=0','1');return false" class="link1">Satýþtaki Itemler</a></strong></td>
	<td align="center"><strong><a href="#" onclick="pageload('Sayfalar/SellingPanel.Asp?Action=MyItems&Durum=1','1');return false" class="link1">Satýlan Itemler</a></strong></td>
	</tr>

	
	</table>
	<%
	Case "MyItems"%><br />
	<table bordercolor="#cccc00" border="1" cellpadding="2" cellspacing="1" align="center" style="border-collapse: collapse;">
	<tr>
	<td align="center" colspan="8"><strong>Satýþ Bilgileri</strong></td>
	</tr>
 	 <tr>	 
 		<td align="center" colspan="2"><strong>Item Adý</strong></td>
		<td align="center"><strong>Adet</strong></td>
		<td align="center"><strong>Satýþ Ücreti</strong></td>
		<td align="center"><strong>Satan Karakter</strong></td>
		<td align="center"><strong>Satýþa Konulma Zamaný</strong></td>
		<%If Not Request.QueryString("Durum")=0 Then
Response.Write("		<td align=""center""><b>Alýcý Karakter</b></td>")
Response.Write("		<td align=""center""><b>Alýþ Zamaný</b></td>")
End If%>
   </tr>
	<%If (Request.QueryString("Durum"))=1 Then
	Set MyItems=Conne.Execute("Select * From SellingItems Where strAccountid='"&Session("username")&"' And Status=1 Or Status=2")
	Else
	Set MyItems=Conne.Execute("Select * From SellingItems Where strAccountid='"&Session("username")&"' And Status=0")
	End If
	Do While Not MyItems.Eof

set det=Conne.Execute("select * from ITEM where num='"&MyItems("num")&"'")
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

iname=server.htmlencode(det("strname"))

if det("strb")=24 or det("stab")=24 or det("dexb")=24 or det("ChaB")=24 or det("intelb")=24 Then
itemname2=replace(iname, "(+0)" , "(+10)" )
else
itemname2=iname
End If

if det("Countable")=1 Then
drtn=MyItems("piece")
elseif det("duration")>0 and det("ItemType")=0 and det("kind") <> 95 Then
drtn=MyItems("durability")
else
drtn=""
End If

itemname=server.htmlencode(itemname2)

itemx="<img width=45 height=45 src=../item/"&resim2(det("num"))&" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&server.htmlencode(itemname)&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"
Response.Write(" 	 <tr>")
Response.Write(" 		<td align=""center""><strong>"&itemx&"</strong></td>")
Response.Write(" 		<td align=""center""><strong>"&itemname&"</strong></td>")
Response.Write("		<td align=""center""><b>"&MyItems("piece")&"</b></td>")
Response.Write("		<td align=""center""><b>"&ayir(MyItems("price"))&"</b></td>")
Response.Write("		<td align=""center""><strong>"&MyItems("strCharId")&"</strong></td>")
Response.Write("		<td align=""center""><b>"&MyItems("SellingTime")&"</b></td>")
If Request.QueryString("Durum")=0 Then
Response.Write("		<td align=""center""><a href=""#"" onclick=""pageload('Sayfalar/SellingPanel.Asp?Action=GeriAl&SellingId="&MyItems("SellingId")&"','1');return false"" class=""link1"">Itemi Satýþtan Çek</a></td>")
Else
Response.Write("		<td align=""center""><b>"&MyItems("BuyerCharId")&"</b></td>")
Response.Write("		<td align=""center""><b>"&MyItems("BuyingTime")&"</b></td>")
End If
Response.Write("   </tr>")
End If
MyItems.MoveNext
Loop%>

</table>
	</td>
	</tr>
</table>
<%Case "TransportBank"
Set MyMoney=Conne.Execute("Select Sum(Price) as ToplamPara From SellingItems Where strAccountid='"&Session("username")&"' And Status=1")
Set BankMoney=Conne.Execute("Select nMoney From Warehouse Where strAccountid='"&Session("username")&"'")
ToplamPara = MyMoney("ToplamPara")
If Isnull(ToplamPara) Then ToplamPara=0
BankadakiPara = BankMoney("nMoney")
If BankadakiPara+ToplamPara>2100000000 Then
Aktarilan=2100000000
Conne.Execute("Update Warehouse set nmoney=2100000000 Where strAccountid='"&Session("username")&"'")
Else
Aktarilan=BankadakiPara+ToplamPara
Conne.Execute("Update Warehouse set nmoney="&Aktarilan&" Where strAccountid='"&Session("username")&"'")
End If
Conne.Execute("Update SellingItems Set Status=2 Where strAccountid='"&Session("username")&"' And Status=1")
Response.Write("<br><br><div class=""errortxt"">Paranýz Bankaya Aktarýlmýþtýr.<br>Bakiye: "&Aktarilan&"</div>")
MyMoney.Close
Set MyMoney=Nothing
Case "GeriAl"
SellingId=Request.QueryString("SellingId")
If len(SellingId)=0 Then Response.End
If isnumeric(SellingId)=False Then Response.End

Set MyItem=Conne.Execute("Select * From SellingItems Where strAccountid='"&Session("username")&"' And SellingId="&SellingId)
If MyItem.Eof Then
Response.Write("<br><br><div class=""errortxt""> Karakter Veya Item Bulunamadý Bulunamadý</div>")
Response.End
End If
Set ong=Conne.Execute("Select * from currentuser where strcharid='"&MyItem("strcharid")&"' ")

If Not ong.Eof Then
Response.Write "<br><br><div class=""errortxt"">Karakteriniz Online Iken Bu Iþlemi Yapamazsýnýz.</div>"
Response.End
End If

Conne.Execute("Exec item_decode2 '"&MyItem("strcharid")&"'")

Set itemkoy = Server.CreateObject("ADODB.Recordset")
SQL="Select * From Inventory_Edit Where Num=0 And struserid='"&MyItem("strcharid")&"' And inventoryslot>13 and inventoryslot<42"
itemkoy.open SQL,conne,1,3

If Not itemkoy.Eof Then
itemkoy("num")=MyItem("num")
itemkoy("stacksize")=MyItem("piece")
itemkoy("durability")=MyItem("durability")
itemkoy("strserial")=MyItem("strserial")
itemkoy.Update

Conne.Execute("Exec Item_encode2 '"&MyItem("strcharid")&"'")
Else
Response.Write("<br><br><div class=""errortxt""> Inventoryde Boþ Yer Olmadýðýndan Item Alýnamadý.<br>Lütfen Inventoryde Boþ Yer Açýn.</div>")
Response.End
End If

Conne.Execute("Delete From Inventory_Edit where struserid='"&MyItem("strcharid")&"'")
Conne.Execute("Delete From SellingItems Where strAccountid='"&Session("username")&"' And SellingId="&SellingId)
Response.Write("<br><br><div class=""errortxt""> Item Inventoryinize Geri Konmuþtur.</div>")
itemkoy.Close
Set itemkoy=Nothing
Case "Buy"%>
<script language="javascript1.3" type="text/javascript">
function BuyingItem(){
$.ajax({
   type: 'post',
   url: 'Sayfalar/SellingPanel.asp?Action=Buying',
   data: $('#chrid').serialize(),
   start:$('div#ortabolum').html('<center><br><br><br><br><img src=imgs/18-1.gif><br><br><font size=1><b>Yükleniyor...</b></font></center>')&&loadingx,
  success: function(ajaxCevap) {
  $('#ortabolum').html(ajaxCevap);
  successx();
   }
});
}
</script>
<style>
.inpt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
border:solid 1px;
border-color:#8E6400;
color:#8E6400;
font-weight:bold;
height:20px;
text-decoration:inherit;
background-color:#F4F4F4
}
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
</style><br />
<br />
<form action="javascript:BuyingItem()" method="post" id="chrid">
<span style="color: rgb(142, 100, 0); font-weight: bold; font-size: 10px;">Karakter Seçiniz: </span><select name="charid" class="inpt" style="padding-top: 2px;">
<%Set Users=Conne.Execute("Select * From Account_Char Where strAccountid='"&session("username")&"'")
If Len(Users("strcharid1"))>0 Then Response.Write "<option value=""1"" style=""height: 15px;"">"&Users("strcharid1")&"</option>"
If Len(Users("strcharid2"))>0 Then Response.Write "<option value=""2"" style=""height: 15px;"">"&Users("strcharid2")&"</option>"
If Len(Users("strcharid3"))>0 Then Response.Write "<option value=""3"" style=""height: 15px;"">"&Users("strcharid3")&"</option>"%>
</select>
<input type="submit" value="Ýleri >>" style="color: rgb(142, 100, 0); font-weight: bold; font-size: 10px;" />
</form>
<%Case "Buying"
charids = secur(Request("charid"))
set charkontrol=Conne.Execute("select * from account_char where straccountid='"&session("username")&"'")

if not charkontrol.eof Then

If charids="1" Then
charid=charkontrol("strcharid1")
ElseIf charids="2" Then
charid=charkontrol("strcharid2")
ElseIf charids="3" Then
charid=charkontrol("strcharid3")
Else
Response.End
End If


%>
<script language="javascript1.2" type="text/javascript">
function satinal(sellingid,itemname,ucret,charid){
if(confirm('<%=Trim(charid)%> Karakterine '+itemname+' Almak Istiyormusunuz?\nKasanýzdan '+ucret+' Coin Eksilecektir.')){
$.ajax({
   type: 'post',
   url: 'Sayfalar/SellingPanel.asp?Action=BuyItem',
   data: 'sellingid='+sellingid+'&alici='+charid,
   start:$('div#ortabolum').html('<center><br><br><br><br><img src=imgs/18-1.gif><br><br><font size=1><b>Yükleniyor...</b></font></center>')&&loadingx,
  success: function(ajaxCevap) {
  $('#ortabolum').html(ajaxCevap);
  successx();
   }
});

}

}
</script> 
<%Set SellingItems=Conne.Execute("Select * From SellingItems Where status=0 Order By SellingId")
Response.Write("<br><table width=""500"">")
Response.Write("<tr><td colspan=""2"" align=""center""><b>Item</b></td><td><b>Adet</b></td><td><b>Ücret</b></td><td><b>Satýcý</b></td><td></td></tr>")
Do While Not SellingItems.Eof

set det=Conne.Execute("select * from ITEM where num='"&SellingItems("num")&"'")
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

iname=server.htmlencode(det("strname"))

if det("strb")=24 or det("stab")=24 or det("dexb")=24 or det("ChaB")=24 or det("intelb")=24 Then
itemname2=replace(iname, "(+0)" , "(+10)" )
else
itemname2=iname
End If

if det("Countable")=1 Then
drtn=SellingItems("piece")
elseif det("duration")>0 and det("ItemType")=0 and det("kind") <> 95 Then
drtn=SellingItems("durability")
else
drtn="1"
End If

itemname=server.htmlencode(itemname2)

itemx="<img width=45 height=45 src=../item/"&resim2(det("num"))&" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&server.htmlencode(itemname)&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"
End If

Response.Write("<tr><td width=""47"">"&itemx&"</td><td>"&itemname&"</td><td>"&drtn&"</td><td>"&ayir(SellingItems("price"))&"</td><td><a href=""#"" onclick=""pageload('Sayfalar/PmBox.Asp?pm=new&userid="&trim(SellingItems("StrCharId"))&"','1');return false"" class=""link1"">"&SellingItems("StrCharId")&"</td><td><a href=""#"" class=""link1"" onclick=""satinal('"&SellingItems("SellingId")&"','"&trim(itemname)&"','"&SellingItems("price")&"','"&charids&"');return false;"">Satýn Al</a></td></tr>")
SellingItems.Movenext
Loop
Response.Write("</table>")
End If%>

<%Case "BuyItem"
sellingid=secur(Trim(Request.Form("sellingid")))
charids=Secur(Trim(Request.Form("alici")))
If isnumeric(charid)=False Or isnumeric(sellingid)=False Then
Response.End
End If

set charkontrol=Conne.Execute("select * from account_char where straccountid='"&session("username")&"'")
If not charkontrol.eof Then
If charids="1" Then
charid=charkontrol("strcharid1")
ElseIf charids="2" Then
charid=charkontrol("strcharid2")
ElseIf charids="3" Then
charid=charkontrol("strcharid3")
Else
Response.End
End If

Set ong=Conne.Execute("Select * from currentuser where strcharid='"&charid&"' ")

If Not ong.Eof Then
Response.Write "<br><br><div class=""errortxt"">Karakteriniz Online Iken Item Satýn Alamazsýnýz.</div>"
Response.End
End If
Set SellingItems=Conne.Execute("Select * From SellingItems Where SellingId="&sellingid)

Set userg = Server.CreateObject("ADODB.Recordset")
SQL="select * from userdata where struserid='"&charid&"'"
userg.open SQL,conne,1,3

If userg("gold")<SellingItems("price") Then
Response.Write "<br><br><div class=""errortxt"">Karakterde Yeterli Para Bulunmamaktadýr.</div>"
Response.End
End if

Conne.Execute("Exec item_decode2 '"&charid&"'")

Set slot = Server.CreateObject("ADODB.Recordset")
SQL="Select * from inventory_edit where num=0 and struserid='"&charid&"' and inventoryslot>13 and inventoryslot<42"
slot.open SQL,conne,1,3

If Not slot.Eof Then

slot("num")=SellingItems("num")
slot("stacksize")=SellingItems("piece")
slot("durability")=SellingItems("durability")
slot("strserial")=SellingItems("strserial")
slot.Update
Conne.Execute("exec item_encode2 '"&charid&"'")
Conne.Execute("Update Userdata set gold=gold-"&SellingItems("price")&" where struserid='"&charid&"' ")
Conne.Execute("Update SellingItems set Status=1,BuyerAccountId='"&session("username")&"',BuyerCharId='"&charid&"',BuyingTime='"&now()&"' Where Sellingid="&sellingid)

Response.Write "<br><br><br><div style=""color: rgb(142, 100, 0);font-weight:bold"">Ýþlem Baþarýlý<br> Item Inventorynizde Boþ Bir Slota Yerleþtirilmiþtir.</div>"
Response.End

'Eðer Inv Doluysa Bankaya Koy
Else

Conne.Execute("exec banka_item_decode '"&session("username")&"'")

Set bslot = Server.CreateObject("ADODB.Recordset")
sql="select * from banka_check where straccountid='"&session("username")&"' and dwid=0"
bslot.Open sql,conne,1,3
If Not bslot.Eof Then
bslot("dwid")=SellingItems("num")
bslot("strserial")=SellingItems("strserial")
bslot("durability")=SellingItems("durability")
bslot("stacksize")=SellingItems("piece")
bslot.Update
Conne.Execute("exec banka_item_encode '"&session("username")&"'")
Conne.Execute("Update Userdata set gold=gold-"&SellingItems("price")&" where struserid='"&charid&"' ")
Conne.Execute("Update SellingItems set Status=1,BuyerAccountId='"&session("username")&"',BuyerCharId='"&charid&"',BuyingTime='"&now()&"' Where Sellingid="&sellingid)
Response.Write "<br><br><br><div style=""color: rgb(142, 100, 0);"">Ýþlem Baþarýlý<br>Item Bankanýzda Boþ Bir Slota Yerleþtirilmiþtir.</div>"
Response.End
Else
Response.Write "<br><br><br><div style=""color: rgb(142, 100, 0);"">HATA!<br>Inventory Ve Bankanýzda Boþ Yer Olmadýðýndan Item Alýnamadý!</div>"
Response.End
End If

End If

Conne.Execute("delete inventory_edit where strUserId='"&userid&"'")
Conne.Execute("delete banka_check where straccountid='"&accid&"'")

End If
Case "Sell"%>
<script language="javascript1.3" type="text/javascript">
function BuyingItem(){
$.ajax({
   type: 'post',
   url: 'Sayfalar/SellingPanel.asp?Action=Selling',
   data: $('#chrid').serialize(),
   start:$('div#ortabolum').html('<center><br><br><br><br><img src=imgs/18-1.gif><br><br><font size=1><b>Yükleniyor...</b></font></center>')&&loadingx,
  success: function(ajaxCevap) {
  $('#ortabolum').html(ajaxCevap);
  successx();
   }
});
}
</script>
<br /><br /><style>
.inpt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
border:solid 1px;
border-color:#8E6400;
color:#8E6400;
font-weight:bold;
height:20px;
text-decoration:inherit;
background-color:#F4F4F4
}
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
</style>
<form action="javascript:BuyingItem()" method="post" id="chrid">
<span style="color: rgb(142, 100, 0); font-weight: bold; font-size: 10px;">Karakter Seçiniz: </span><select name="charid" class="inpt" style="padding-top: 2px;">
<%Set Users=Conne.Execute("Select * From Account_Char Where strAccountid='"&session("username")&"'")
If Len(Users("strcharid1"))>0 Then Response.Write "<option value=""1"" style=""height: 15px;"">"&Users("strcharid1")&"</option>"
If Len(Users("strcharid2"))>0 Then Response.Write "<option value=""2"" style=""height: 15px;"">"&Users("strcharid2")&"</option>"
If Len(Users("strcharid3"))>0 Then Response.Write "<option value=""3"" style=""height: 15px;"">"&Users("strcharid3")&"</option>"%>
</select>
<input type="submit" value="Ýleri >>" style="color: rgb(142, 100, 0); font-weight: bold; font-size: 10px;" />
</form>
<%Case "Selling"
charids = Request("charid")
set charkontrol=Conne.Execute("select * from account_char where straccountid='"&session("username")&"'")

if not charkontrol.eof Then

If charids="1" Then
charid=charkontrol("strcharid1")
ElseIf charids="2" Then
charid=charkontrol("strcharid2")
ElseIf charids="3" Then
charid=charkontrol("strcharid3")
Else
Response.End
End If
%>
<script>
function selectedbox(slot){
 for (var i=14; i<42; i++){
  var spn = document.getElementById(i);
  spn.style.border = "1px solid #4C4B36";
	 }
	document.getElementById(slot).style.border = "1px solid #00FF00";
	document.getElementById('inventoryslot').value=slot;

	document.getElementById('moneys').style.visibility="visible";
	document.getElementById('satis').style.visibility="visible";
	}
function karakterfiltre(deger) {
var yaz = new String();
var numaralar = "0123456789";
var chars = deger.value.split("");
for (i = 0; i < chars.length; i++) {
if (numaralar.indexOf(chars[i]) != -1) yaz += chars[i];
}
if (deger.value != yaz) deger.value = yaz;
}

function solditem(){
if ($('#money').val()==''){
alert('Satýþ Ücretini Boþ Býrakmayýnýz.')
}
else{
$.ajax({
   type: 'post',
   url: 'Sayfalar/SellingPanel.asp?Action=Sold',
   data: $('#solditemm').serialize(),
   start:$('div#ortabolum').html('<center><br><br><br><br><img src=imgs/18-1.gif><br><br><font size=1><b>Yükleniyor...</b></font></center>')&&loadingx,
   success: function(ajaxCevap) {
  $('#ortabolum').html(ajaxCevap);
  successx();
   }
});
}
}
</script>
<style>
.inpt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
border:solid 1px;
border-color:#8E6400;
color:#8E6400;
font-weight:bold;
height:20px;
text-decoration:inherit;
background-color:#F4F4F4
}
.txt{
font-family:Verdana, Arial, Helvetica, sans-serif;
font-size:10px;
color:#8E6400;
font-weight:bold;
}
</style><br />
<br />
<form action="javascript:solditem();" method="post" id="solditemm" name="solditemm">
<%
Dim items(42)
Dim adet(42)
Set userdt = Conne.Execute("SELECT * FROM USERDATA WHERE strUserId='"&charid&"'")
If userdt.Eof Then
Response.Write("<font color=""#FF0000""><strong>Karakter Bulunamadý</strong></font>")
Response.End
End If

Set onlinek=Conne.Execute("select strcharid from currentuser where strcharid='"&charid&"'")
If Not onlinek.Eof Then
Response.Write("Karakter Online Iken Item Satýþa Konulamaz.")
Response.End
End If

conne.Execute("Truncate table inventory_edit")
conne.Execute("Exec item_decode2 '"&charid&"'")
Dim itemler,userdt
set itemler=Conne.Execute("SELECT * FROM inventory_edit WHERE struserid='"&charid&"' And inventoryslot>13 and inventoryslot<42")
%>
<table border="0" align="center" bordercolor="333333" cellpadding="0" cellspacing="0" style="color:#FFFFFF;">
            <tr>
			<td width="350" height="600" valign="top">
<div align="left" id="inven" style="color:#FFFFFF;font-size:11px;font-family:Arial,Verdana,Helvetica, sans-serif;font-weight:bold;background:url(../imgs/inv.gif);background-repeat:no-repeat; height:100%;width:366; overflow:hidden">
<%If not itemler.eof Then
Dim x,itemi,det,dtype,speed,kinds,delay,kind,renk,atack,weight,durability,duration,defans,dodging,incap,daggerac,swordac,clubac,axeac,spearac,bowac,firedam,icedam,ligthdam,posdam,hpdrain,mpdamage,mpdrain,mirrordam,strbon,canbonus,hpbon,dexbon,intbon,mpbon,magicbon,fireres,glares,lightres,magicres,posres,curseres,ReqStr,reqhp,reqdex,Reqint,Reqcha,iname,itemname2,drtn,sid,itemname,slot,maceac,xx
for x=14 to 41
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
items(slot)="<img width=45 height=45 src=../item/"&resim2(det("num"))&" onMouseOver=""return overlib('<body bgcolor=#000000><b><center><font style=font-size:11px color="&renk&">"&server.htmlencode(itemname)&"<br>"&dtype&"</font><br><font color=white style=font-size:11px>"&kind&"</font><br><br></center><font color=white style=font-size:11px;>"&atack&delay&weight&duration&defans&dodging&incap&"</font><font color=lime style=font-size:11px>"&daggerac&swordac&maceac&axeAc&spearac&bowac&firedam&icedam&ligthdam&posdam&hpdrain&mpdamage&mpdrain&mirrordam&strbon&canbonus&hpbon&dexbon&intbon&mpbon&magicbon&fireres&glares&lightres&magicres&posres&curseres&"</font><font color=white style=font-size:11px>"&ReqStr&reqhp&reqdex&Reqint&Reqcha&"</font>', RIGHT, WIDTH, 240,CELLPAD, 5, 10, 10);"" onMouseOut=""return nd();"">"

End If

itemler.movenext
next

for xx=15 to 42
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
<table width="366" height="166" border="0" bordercolor="#333333" cellpadding="1" cellspacing="2" style="position:relative;left:0px;top:9px;">
  <tr >
    <td width="45" height="45" id="14" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(15)%></td>
    <td width="45" height="45" id="15" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(16)%></td>
    <td width="45" height="45" id="16" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(17)%></td>
    <td width="45" height="45" id="17" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(18)%></td>
    <td width="45" height="45" id="18" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(19)%></td>
    <td width="45" height="45" id="19" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(20)%></td>
    <td width="45" height="45" id="20" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(21)%></td>
  </tr>
  <tr >
    <td  width="45" height="45" id="21" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(22)%></td>
    <td width="45" height="45" id="22" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(23)%></td>
    <td width="45" height="45" id="23" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(24)%></td>
    <td width="45" height="45" id="24" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(25)%></td>
    <td width="45" height="45" id="25" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(26)%></td>
    <td width="45" height="45" id="26" style="border:1px solid #4C4B36"  onClick="selectedbox(this.id);"><%=items(27)%></td>
    <td width="45" height="45" id="27" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(28)%></td>
  </tr>
  <tr>
    <td width="45" height="45" id="28" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(29)%></td>
    <td width="45" height="45" id="29" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(30)%></td>
    <td width="45" height="45" id="30" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(31)%></td>
    <td width="45" height="45" id="31" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(32)%></td>
    <td width="45" height="45" id="32" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(33)%></td>
    <td width="45" height="45" id="33" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(34)%></td>
    <td width="45" height="45" id="34" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(35)%></td>
  </tr>
  <tr>
    <td width="45" height="45" id="35" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(36)%></td>
    <td width="45" height="45" id="36" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(37)%></td>
    <td width="45" height="45" id="37" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(38)%></td>
    <td width="45" height="45" id="38" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(39)%></td>
    <td width="45" height="45" id="39" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(40)%></td>
    <td width="45" height="45" id="40" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(41)%></td>
    <td width="45" height="45" id="41" style="border:1px solid #4C4B36" onClick="selectedbox(this.id);"><%=items(42)%></td>
  </tr>
</table><style>.adt{height:14px;}</style>
<%topx=(-149)
leftx=-46
For aa=1 To 28
If aa=8 or aa=15 or aa=22 Then
topx=topx+51-14
leftx=-46+52
Else
topx=(topx)-14
leftx=(leftx)+52
End If%><div align="left" class="adt" id="adet<%=aa+13%>" style="position:relative; top:<%=topx%>px; left:<%=leftx%>px; width:20px"><%=adet(aa+14)%></div>
<%next%>
<%
ElseIf charid="" Then
Response.End
Else
Response.Write "Kullanýcý Bulunamadý!"
End If%>
</div>
</td>
<td valign="top" style="padding-left:20px">
<table>
<%Response.Write "<input type=""hidden"" value="""&charids&""" name=""charid"">"&vbcrlf
Response.Write "<input type=""hidden"" value=""14"" name=""inventoryslot"" id=""inventoryslot"">"&vbcrlf
%>
<tr><td><div id="moneys"  style="visibility:hidden;color:#8E6400; font-weight:bold">Satýþ Ücreti: <input type="text" dir="rtl" name="money" id="money" class="inpt" onkeyup="karakterfiltre(this);"/></div></td></tr></table>
<input type="submit" value="Satýþa Çýkar" style="color: rgb(142, 100, 0); font-weight: bold; font-size: 10px; visibility:hidden" id="satis" /><br /><br />
</td>
</tr>
</table>
</form>
<%Conne.Execute("Delete From Inventory_Edit where strUserId='"&charid&"'")
End If %>


<%Case "Sold"
charids=Secur(Request("charid"))
inventoryslot=Secur(Request.Form("inventoryslot"))
money=Secur(Request.Form("money"))
If isnumeric(money)=False Then
Response.Write("<br><b>Satýþ Fiyatý Bölümüne Sadece Sayýsal Deðer Giriniz.</b>")
Response.End
End If

If int(money)>2100000000 Then
Response.Write("<br><br><div class=""errortxt"">Itemi Max 21 GB ye Satabilirsiniz.</div>")
Response.End
End If
set charkontrol=Conne.Execute("select * from account_char where straccountid='"&session("username")&"'")

if not charkontrol.eof Then

If charids="1" Then
charid=charkontrol("strcharid1")
ElseIf charids="2" Then
charid=charkontrol("strcharid2")
ElseIf charids="3" Then
charid=charkontrol("strcharid3")
Else
Response.End
End If

Set onlinek=Conne.Execute("select strcharid from currentuser where strcharid='"&charid&"'")

If Not onlinek.Eof Then
Response.Write("<br><br><div class=""errortxt"">Karakter Online Iken Item Satýþa Konulamaz.</div>")
Response.End
End If
conne.Execute("truncate table inventory_edit")
conne.Execute("Exec item_decode2 '"&charid&"'")
Dim itemlr
Set itemlr=Conne.Execute("SELECT * FROM inventory_edit WHERE struserid='"&charid&"' And inventoryslot="&inventoryslot)
If itemlr.Eof Then
Response.Write("<br><br><div class=""errortxt"">Geçerli Bir Item Satýþa Koyun.</div>")
Response.End
End If

If left(itemlr("num"),1)=9 Then
Response.Write("<br><br><div class=""errortxt"">Bu Item Satýþa Konulamaz.</div>")
Response.End
End If

Set PremiumItemKontrol=Conne.Execute("Select premiumitem From PUS_ITEMLERI Where itemkodu="&itemlr("num")&" And premiumitem=1")
If Not PremiumItemKontrol.Eof Then
Response.Write("<br><b>Bu Item Satýþa Konulamaz.</b>")
Response.End
End If

Set det=Conne.Execute("Select * From ITEM Where Num='"&itemlr("num")&"'")
If Not det.Eof Then

If det("Countable")=1 Then
adet=itemlr("stacksize")
ElseIf det("duration")>0 And det("ItemType")=0 And det("kind") <> 95 Then
adet=itemlr("durability")
Else
adet="1"
End If

Conne.Execute("insert into SellingItems(StrAccountId,StrCharId,Num,Durability,Piece,StrSerial,ItemType,ItemClass,Price,SellingTime) values('"&session("username")&"','"&charid&"',"&itemlr("num")&","&itemlr("durability")&","&itemlr("stacksize")&","&itemlr("StrSerial")&","&det("ItemType")&","&det("Class")&","&money&",'"&Now()&"')")
Conne.Execute("Update Inventory_Edit Set Num=0, stacksize=0, Durability=0, strserial=0 where strUserId='"&charid&"' and inventoryslot='"&inventoryslot&"'")
Conne.Execute("Exec Item_encode2 '"&charid&"'")
Conne.Execute("Delete From Inventory_Edit where strUserId='"&charid&"'")
Response.Write("<br><b>Item Satýþa Konuldu.<br>Iteminiz Satýlýrsa Paranýzý Satýþ Paneli Bölümünden Alabilirsiniz.")
End If
End If
End Select


Else
Response.Write "<div class=""errortxt"">Lütfen kullanýcý giriþi yapýnýz.</div>"
End If
Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>