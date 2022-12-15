<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<% Dim MenuAyar,ksira,accid,itemid,itemara,pusitem,cashp
if Session("durum")="game" and Session("accountid")<>"" Then
accid=Session("accountid")

elseif Session("login")="ok" and Session("username")<>"" Then
accid=Session("username")

else
server.execute("hata.asp")
Response.End
End If

Response.expires=0
Response.Charset = "iso-8859-9"

Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then
itemid=secur(Request.Querystring("itemid"))
if isnumeric(itemid)=false Then
Response.Write "Item Bulunamadi!"
Response.End
End If
set itemara=Conne.Execute("select num from item where num="&itemid&"")
if itemara.eof Then
Response.Write "<br><br><font color=""white""><center>Item Bulunamadi!</center></font>"
Response.End
End If

set pusitem=Conne.Execute("select * from pus_itemleri where itemkodu="&itemid&"")
set cashp=Conne.Execute("select * from tb_user where straccountid='"&accid&"'")

if not pusitem.eof Then %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head>
<title>Power Up Store</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-2">
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

function buysend(num,charid){
	location.href="buyitem.asp?itemid="+num+"&charid="+charid;
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
          <td style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" align="center">* <strong>Item Satýn Al</strong></td>
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
			  
              <!--  šYþš lA				 -->
              <table border="0" cellpadding="0" cellspacing="0" width="215">

                <tbody><tr>
                  <td height="20"></td>
                </tr>
                <tr>
                  <td><table border="0" cellpadding="0" cellspacing="0" width="215">
                      <tbody>
	<%if pusitem("premiumitem")=1 Then%>
                      <tr>
		<td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" width="110" height="21">* Kullaným Günü: </td>
		<td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" align="center"><%=pusitem("kullanimgunu")%> Gün</td>
                      </tr>
		<%End If%>
		<tr>
                        <td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" width="110" height="21">*
                          Ücret <br> </td>
                        <td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" align="center"><%=pusitem("ucret")%></td>

                      </tr>
                      <tr>
                        <td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" height="21">*
                          Bakiye</td>
                        <td class="12_white" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" align="center"><%=cashp("cashpoint")%></td>
                      </tr>
                      <tr>
                        <td class="12_yellow140" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" height="21">*
                          Gelecek Bakiye</td>

                        <td class="12_yellow140" style="padding-left: 8px; font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" align="center"><%=cashp("cashpoint")-pusitem("ucret")%></td>
                      </tr>
                    </tbody></table></td>
                </tr>
              </tbody></table>
              <table border="0" cellpadding="0" cellspacing="0" width="96%" height="40">
                <tbody>
	<form action="buyitem.asp" method="get" name="buy" id="buy">
	<input type="hidden" value="<%=pusitem("itemkodu")%>" name="itemid">
	<tr><td>&nbsp;</td></tr>
<% if Session("durum")="web" Then %>
	<tr align="center" valign="bottom">
		<td>
<font style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;">Karakterinizi Seçiniz: </font></td>
	<td>
	<select name="charid" id="charid" style="background-color:#000000;font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;">
	<%dim accch
	Set accch =Conne.Execute("Select * From ACCOUNT_CHAR where strAccountID='"&accid&"'")
	if not accch.eof Then 
	dim charid1,charid2,charid3
	charid1=trim(accch("strcharid1"))
	charid2=trim(accch("strcharid2"))
	charid3=trim(accch("strcharid3"))
	End If
if len(charid1)>0 Then
Response.Write "<option value=""1"">"&charid1&"</option>"
End If
if len(charid2)>0 Then
Response.Write "<option value=""2"">"&charid2&"</option>"
End If
if len(charid3)>0 Then
Response.Write "<option value=""3"">"&charid3&"</option>"
End If
%>
</select>
</td></tr></form><%End If%><tr><td>&nbsp;</td></tr>
		<tr align="center" valign="bottom">
                  <td><a href="javascript:document.getElementById('buy').submit();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_sattle','','images/butt/butt_settle_on.gif',1)"><img src="images/butt/butt_settle.gif" name="butt_sattle" border="0" width="89" height="30"></a></td>

                  <td><a href="javascript:wish_go(<%=pusitem("itemkodu")%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_cart111','','images/butt/butt_cart_on.gif',1)"><img src="images/butt/butt_cart.gif" name="butt_cart111" id="butt_cart1" border="0" width="89" height="30"></a></td>
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

</td></tr></tbody></table>
    </td>
  </tr>
  <tr>
    <td align="center" height="44">
<script language="JavaScript" type="text/JavaScript">

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
Response.Write "Item bulunamadi !"
End If

else
Response.Write "<br><b><span style=""color:#000000;font-family:Verdana, Arial, Helvetica,sans-serif;font-size:10px;"">Bu bölüm Admin tarafindan kapatilmiºtir.</span></b>"
End If
MenuAyar.Close
Set MenuAyar=Nothing%>