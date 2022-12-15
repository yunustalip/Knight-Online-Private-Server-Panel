<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%Dim MenuAyar,ksira,itemid,charid,pusitem,cashp,itemoz,accid,sql,charkontrol,serial,ong,slot,insertprivate

Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then

If Session("durum")="game" and Session("accountid")<>"" Then
accid=Session("accountid")
charid=Session("charid")
ElseIf Session("durum")="web" and Session("login")="ok" and Session("username")<>"" and secur(Request.Querystring("charid"))<>"" Then
accid=Session("username")
charid=secur(Request.Querystring("charid"))
Else
Server.Execute("hata.asp")
Response.End
End If


itemid=secur(Request.Querystring("itemid"))
If IsNumeric(itemid)=False Then
Response.End 
End If

Set pusitem = Server.CreateObject("ADODB.Recordset")
sql="select * from pus_itemleri where itemkodu="&itemid&""
pusitem.open sql,conne,1,3

Set cashp = Server.CreateObject("ADODB.Recordset")
sql=("Select * From tb_user where straccountid='"&accid&"'")
cashp.open sql,conne,1,3

If pusitem.eof Then
Response.Redirect("hit.asp")
Response.End 
End If 

If Session("durum")="web" and Session("username")<>"" Then
set charkontrol=Conne.Execute("select * from account_char where straccountid='"&accid&"'")

if not charkontrol.eof Then

if charid="1" Then
charid=charkontrol("strcharid1")
elseif charid="2" Then
charid=charkontrol("strcharid2")
elseif charid="3" Then
charid=charkontrol("strcharid3")
else
Response.End
End If

else
Response.End 
End If
else
charid=Session("charid")
End If

If cashp.eof Then
Response.End
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head>
<title>Power Up Store</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9">
<link href="images/style.css" rel="stylesheet" type="text/css">
<script language="Javascript">

<!--

function detail_go(num) {
	location.href="detay.asp?itemid="+num;
}
//-->
</script>
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

function buysend(num){
	location.href="buyitem.asp?itemid="+num;
}
function wish_go(num) {
	location.href="cart.asp?itemid="+num;
}
//-->
</script>
<style type="text/css">
body {background-color: transparent}
</style>
</head><body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="MM_preloadImages('/images/butt/butt_cart_cencle_on.gif','/images/butt/butt_settle_on.gif','/images/butt/butt_all_cencle_on.gif','/images/butt/butt_cart_on.gif')">

<table border="0" cellpadding="0" cellspacing="0" width="288">
  <tbody><tr>
    <td height="31">


<table style="position: relative; top: -5px;" border="0" cellpadding="0" cellspacing="0">
  <tbody><tr>
        <td width="8"></td>
		<td style="padding-top: 0px;">
			<a href="cart.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_cart','','images/right_sub_memu_02.gif',1)">
			<img src="images/right_sub_memu_02.gif" name="butt_cart" border="0" width="100" height="29">			</a>		</td>
        <td width="1"></td>
	    <td style="padding-top: 0px;">
		<img src="images/right_sub_memu_03.gif" name="butt_detail" border="0" width="59" height="29">		</td>

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
          <td style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" align="center">* <strong>Item Satin Al</strong></td>
        </tr>
        <tr>

          <td class="12_white" align="center" valign="bottom"> <div id="item_list" style="border: 0px none ; overflow: auto; height: 400px;">
              <table border="0" cellpadding="0" cellspacing="0" width="212">
                <tbody><tr>
                  <td height="10"></td>
                  <td></td>
                </tr>
                <tr>

                  <td valign="bottom" width="83"><table border="0" cellpadding="0" cellspacing="0" width="74">
                      <tbody><tr>
                        <td width="72"><table border="0" cellpadding="0" cellspacing="0" width="72" height="72">
                            <tbody><tr>
                              <td align="center" background="images/table1_01.gif"><a href="javascript:detail_go(<%=itemid%>);"><img src="item/<%=resim(pusitem("resim"))%>" width="64" height="64" border="0"></a><a href="javascript:detail_go(<%=itemid%>);"></a></td>
                            </tr>
                          </tbody></table></td>
                      </tr>
                    </tbody></table></td>

                  <td><table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tbody><tr>
                        <td style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" height="28">*<strong> <%=pusitem("itemismi")%> </strong></td>
                      </tr>
                      <tr>
                        <td style="font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;">* Adet : <%=pusitem("adet")%> </td>
                      </tr>
                    </tbody></table></td>
                </tr>
              </tbody></table>
			  
              <table border="0" cellpadding="0" cellspacing="0" width="215">

                <tbody><tr>
                  <td height="20"></td>
                </tr>
                <tr>
                  <td><table border="0" cellpadding="0" cellspacing="0" width="215">
                      <tbody>
                      <tr>
                        <td width="110" height="21" class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;">*
                          Bakiye</td>
                        <td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" align="center"><%=cashp("cashpoint")%></td>
                      </tr>
                    </tbody></table></td>
                </tr>
              </tbody></table>
              <table border="0" cellpadding="0" cellspacing="0" width="96%" height="40">
                <tbody><tr align="center" valign="bottom">
                  <td>

<%set itemoz=Conne.Execute("select * from item where num="&itemid&"")
if itemoz.eof Then
Response.Write "Item Bulunamadý"
Response.End
End If

if cashp("cashpoint")<pusitem("ucret") Then
Response.Write "<font color=#ffee2f face=ms sans serif><br><br><center>Yetersiz Cash Point.<br>En az "&pusitem("ucret")&" Cash pointinizin bulunmasý gereklidir.<br><br>Cash Point satýn almak için <a href='powerupstore.asp?control=recharge'><font color=#ffee2f face=ms sans serif>týklayýn</font></a></font>"
Response.End
Else

'Hit Arttýr
If Not pusitem.Eof Then
pusitem("alindi")=pusitem("alindi")+1
pusitem.Update

'Premium Itemler Ýçin Baþlangýç
If Pusitem("premiumitem")=1 Then
Randomize
serial=int((rnd*899999999)+100000000)

Set ong=Conne.Execute("select * from currentuser where strcharid='"&charid&"' ")

If Not ong.Eof Then
Response.Write "<font color=#c3cccc face=ms sans serif><br><br><center>HATA!</font><br><font color=""#ffee2f"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px""> Bu Itemi Alabilmek Ýçin Oyundan Çýkýp Web Siteden <br> Almanýz Gerekmektedir.</font></center>"
Response.End
End If

Conne.Execute("Exec item_decode2 '"&charid&"'")

Set slot = Server.CreateObject("ADODB.Recordset")
SQL="select * from inventory_edit where num=0 and struserid='"&charid&"' and inventoryslot>13 and inventoryslot<42"
slot.open SQL,conne,1,3

If Not slot.Eof Then

slot("num")=itemid
slot("stacksize")=1
slot("durability")=itemoz("duration")
slot("strserial")=serial
slot.Update
Conne.Execute("exec item_encode2 '"&charid&"'")
cashp("cashpoint")=cashp("cashpoint")-pusitem("ucret")
cashp.Update

Set insertprivate = Server.CreateObject("ADODB.Recordset")
sql="select * from Privateitems"
insertprivate.Open sql,conne,1,3
insertprivate.AddNew
insertprivate(0)=itemid
insertprivate(1)=serial
insertprivate(2)=accid
insertprivate(3)=Now
insertprivate(4)=pusitem("kullanimgunu")
insertprivate(5)="alýndý"
insertprivate.Update
insertprivate.Close
Set insertprivate=Nothing
Response.Write "<font color=#c3cccc face=ms sans serif><br><br><center>Ýþlem Baþarýlý</font><br><font color=#ffee2f face=ms sans serif> Item Inventorynizde Boþ Bir Slota Yerleþtirilmiþtir.</font></center>"
Response.End

'Eðer Inv Doluysa Bankaya Koy
Else

Conne.Execute("exec banka_item_decode '"&accid&"'")

Set bslot = Server.CreateObject("ADODB.Recordset")
sql="select * from banka_check where straccountid='"&accid&"' and dwid=0"
bslot.Open sql,conne,1,3
If Not bslot.Eof Then
bslot("dwid")=itemid
bslot("strserial")=serial
bslot("durability")=itemoz("duration")
bslot("stacksize")=1
bslot.Update
Conne.Execute("exec banka_item_encode '"&accid&"'")
Response.Write "<font color=#c3cccc face=ms sans serif><br><br><center>Ýþlem Baþarýlý</font><br><font color=#ffee2f face=ms sans serif> Item Bankanýzda Boþ Bir Slota Yerleþtirilmiþtir.</font></center>"
cashp("cashpoint")=cashp("cashpoint")-pusitem("ucret")
cashp.update

Set insertprivate = Server.CreateObject("ADODB.Recordset")
sql="select * from Privateitems"
insertprivate.open sql,conne,1,3
insertprivate.addnew
insertprivate(0)=itemid
insertprivate(1)=serial
insertprivate(2)=accid
insertprivate(3)=now
insertprivate(4)=pusitem("kullanimgunu")
insertprivate.update
insertprivate.close
Set insertprivate=nothing
Response.End
Else
Response.Write "<font color=#c3cccc face=ms sans serif><br><br><center>HATA!</font><br><font color=#ffee2f face=ms sans serif> Inventory Ve Bankanýzda Boþ Yer Olmadýðýndan Item Alýnamadý!</font></center>"
Response.End
End If

End If

Conne.Execute("delete inventory_edit where strUserId='"&userid&"'")
Conne.Execute("delete banka_check where straccountid='"&accid&"'")

'Premium Item Deðilse
Else

'Web den Giriþ Yaptý ise
If Session("durum")="web" and Session("login")="ok" and Session("username")<>"" and secur(Request.Querystring("charid"))<>"" Then 

Set ong=Conne.Execute("select * from currentuser where strcharid='"&charid&"' ")
If ong.Eof Then

Conne.Execute("Exec item_decode2 '"&charid&"'")

Set slot = Server.CreateObject("ADODB.Recordset")
SQL="Select * From Inventory_Edit Where Num=0 And struserid='"&charid&"' And inventoryslot>13 and inventoryslot<42"
slot.open SQL,conne,1,3

'Inventorye Koy
If Not slot.Eof Then

slot("num")=itemid
slot("stacksize")=1
slot("durability")=itemoz("duration")
slot("strserial")=serial
slot.Update
Conne.Execute("exec item_encode2 '"&charid&"'")
cashp("cashpoint")=cashp("cashpoint")-pusitem("ucret")
cashp.Update

Response.Write "<font color=""#c3cccc"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px""><br><br><center>Ýþlem Baþarýlý</font><br><font color=#ffee2f face=ms sans serif> Item Inventorynizde Boþ Bir Slota Yerleþtirilmiþtir.</font></center>"
Response.End

'Eðer Inventory Doluysa Bankaya Koy
Else

Conne.Execute("exec banka_item_decode '"&accid&"'")

Set bslot = Server.CreateObject("ADODB.Recordset")
sql="Select * from banka_check where straccountid='"&accid&"' and dwid=0"
bslot.open sql,conne,1,3
If Not bslot.Eof Then
bslot("dwid")=itemid
bslot("strserial")=serial
bslot("durability")=itemoz("duration")
bslot("stacksize")=1
bslot.Update
Conne.Execute("exec banka_item_encode '"&accid&"'")
Response.Write "<font color=#c3cccc face=ms sans serif><br><br><center>Ýþlem Baþarýlý</font><br><font color=#ffee2f face=ms sans serif> Item Bankanýzda Boþ Bir Slota Yerleþtirilmiþtir.</font></center>"
cashp("cashpoint")=cashp("cashpoint")-pusitem("ucret")
cashp.Update
Response.End
Else
Response.Write "<font color=""#c3cccc"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px""><br><br><center>HATA!</font><br><font color=#ffee2f face=ms sans serif> Inventory Ve Bankanýzda Boþ Yer Olmadýðýndan Item Alýnamadý!</font></center>"
Response.End
End If
End If

Dim buy2
Set buy2 = Server.CreateObject("ADODB.Recordset")
sql = "Select * From WEB_ITEMMALL_LOG"
buy2.Open sql,conne,1,3
buy2.AddNew
buy2("straccountid")=accid
buy2("strcharid")=charid
buy2("serverno")=1
buy2("itemid")=itemid
buy2("itemcount")=pusitem("adet")
buy2("buytime")=Now
buy2("price")=pusitem("ucret")
buy2("img_file_name")=pusitem("resim")
buy2("strItemName")=pusitem("itemismi")
buy2.Update
buy2.Close
Set buy2=nothing

'Web den Giriþ Yaptý Ve Karakter Oyunda Ise
Else
Dim buy
Set buy = Server.CreateObject("ADODB.Recordset")
sql = "Select * From WEB_ITEMMALL"
buy.open sql,conne,1,3
buy.addnew
buy("straccountid")=accid
buy("strcharid")=charid
buy("serverno")=1
buy("itemid")=itemid
buy("itemcount")=pusitem("adet")
buy("buytime")=now
buy("price")=pusitem("ucret")
buy("img_file_name")=pusitem("resim")
buy("strItemName")=pusitem("itemismi")
buy.update
buy.close
set buy=nothing

Dim buy3
Set buy3 = Server.CreateObject("ADODB.Recordset")
sql = "Select * From WEB_ITEMMALL_LOG"
buy3.Open sql,conne,1,3
buy3.AddNew
buy3("straccountid")=accid
buy3("strcharid")=charid
buy3("serverno")=1
buy3("itemid")=itemid
buy3("itemcount")=pusitem("adet")
buy3("buytime")=now
buy3("price")=pusitem("ucret")
buy3("img_file_name")=pusitem("resim")
buy3("strItemName")=pusitem("itemismi")
buy3.update
buy3.close
Set buy3=Nothing
Response.Write "<font color=""#c3cccc"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px""><br><br><center>Ýþlem Baþarýlý</font><br><font color=#ffee2f style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:12px"">Item Alýndý!<br> Oyuna Girince,<br>Power Up Storeyi 1 Kez Açýp Kapatarak Itemi Inventorynize alabilirsiniz.</font></center>"
Response.End
End If
'Oyundan Giriþ Ise
ElseIf Session("durum")="game" Then

Dim buyongame
Set buyongame = Server.CreateObject("ADODB.Recordset")
sql = "Select * From WEB_ITEMMALL"
buyongame.open sql,conne,1,3
buyongame.addnew
buyongame("straccountid")=accid
buyongame("strcharid")=charid
buyongame("serverno")=1
buyongame("itemid")=itemid
buyongame("itemcount")=pusitem("adet")
buyongame("buytime")=now
buyongame("price")=pusitem("ucret")
buyongame("img_file_name")=pusitem("resim")
buyongame("strItemName")=pusitem("itemismi")
buyongame.update
buyongame.close
set buyongame=nothing

Dim buyongame2
Set buyongame2 = Server.CreateObject("ADODB.Recordset")
sql = "Select * From WEB_ITEMMALL_LOG"
buyongame2.open sql,conne,1,3
buyongame2.addnew
buyongame2("straccountid")=accid
buyongame2("strcharid")=charid
buyongame2("serverno")=1
buyongame2("itemid")=itemid
buyongame2("itemcount")=pusitem("adet")
buyongame2("buytime")=now
buyongame2("price")=pusitem("ucret")
buyongame2("img_file_name")=pusitem("resim")
buyongame2("strItemName")=pusitem("itemismi")
buyongame2.update
buyongame2.close
Set buyongame2=Nothing
Response.Write "<font color=#c3cccc face=ms sans serif><br><br><center>Ýþlem Baþarýlý</font><br><font color=#ffee2f face=ms sans serif> Item Alýndý. </font></center>"

End If
'Pre Item Deðilse bitiþ
End If



Else
Response.Write "Item Bulunamadý."
End If
End If
%></td>
              </tr>
            </tbody></table>
        </div></td>
       </tr>
      <tr>
         <td class="12_white" align="center">&nbsp;</td>
      </tr>
   </tbody></table>
     </td>
  </tr>
  <tr>
      <td style="padding-top: 10px;" align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="252" height="26"><tbody><tr><td style="background-image: url(images/n_table2_02.gif); background-repeat: no-repeat;" height="26">
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr>
  <td style="font-size: 11px; color: rgb(255, 238, 47); font-weight: bold;" align="right"><strong><%=cashp("cashpoint")%> P</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr></tbody></table>
</td></tr>
</tbody>
</table>
    </td>
  </tr>
  <tr>
    <td align="center" height="44">
<table border="0" cellpadding="0" cellspacing="0" width="254">
  <tbody><tr> 
    <td align="center"><a href="my_item.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_item','','images/butt/butt_item_on.gif',1)"><img src="images/butt/butt_item.gif" name="butt_item" border="0" width="119" height="30"></a></td>
  </tr>
</tbody></table>
    </td>
  </tr>
</tbody></table>
</body></html>
<%Else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafýndan kapatýlmýþtýr.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>