<!--#include file="_inc/conn.asp"-->
<!--#include file="function.asp"-->
<%
if Session("durum")="game" and Session("accountid")<>"" Then
accid=Session("accountid")

elseif Session("login")="ok" and Session("username")<>"" Then
accid=Session("username")

else
server.execute("hata.asp")
Response.End
End If
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head>
<title>Hit Item</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-2">
<link href="images/style.css" rel="stylesheet" type="text/css">
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
function wish_go(num) {
	location.href="cart.asp?itemid="+num;
}
function detail_go(num) {
	location.href="detay.asp?itemid="+num;
}
//-->
</script>
<style type="text/css">
body {background-color: transparent}
</style>
</head><body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="MM_preloadImages('/butt/butt_cart_cencle_on.gif','/butt/butt_all_cencle_on.gif','/butt/butt_cart_on.gif','/butt/butt_detail_on.gif','/butt/butt_buy_on.gif','images/right_sub_memu_02.gif')">
<table border="0" cellpadding="0" cellspacing="0" width="288">
  <tbody><tr> 
    <td valign="bottom" height="30">
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
<%Dim MenuAyar,ksira,accid,itype,cashp,hit,k
Set MenuAyar=conne.Execute("select PSt from MenuAyar Where PId='PowerUpStore'")
If MenuAyar("PSt")=1 Then

itype=secur(Request.Querystring("type"))
set cashp=Conne.Execute("select * from tb_user where straccountid='"&accid&"'")
set hit=Conne.Execute("Select TOP 3 * from pus_itemleri where type='"&itype&"' order by alindi desc")
if hit.eof Then
set hit=Conne.Execute("Select TOP 3 * from pus_itemleri order by alindi desc")
End If%>
<table style="position: relative; top: -5px;" border="0" cellpadding="0" cellspacing="0">
  <tbody><tr>
        <td width="8"></td>
		<td style="padding-top: 0px;"><a href="cart.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_cart','','images/right_sub_memu_02b.gif',1)"><img src="images/right_sub_memu_02.gif" name="butt_cart" border="0" width="100" height="29"></a></td>
        <td width="1"></td>
		<td style="padding-top: 0px;"><a href="hit.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('butt_detail','','images/right_sub_memu_03b.gif',1)"><img src="images/right_sub_memu_03b.gif" name="butt_detail" border="0" width="59" height="29"></a></td>
		<td width="1"></td>

    </tr>
</tbody></table>

    </td>
  </tr>
  <tr> 
    <td align="center" valign="top" width="288" height="473"> 
      <table border="0" cellpadding="0" cellspacing="0">
        <tbody><tr> 
          <td height="28">&nbsp;</td>
        </tr>
        <tr> 
          <td style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px; padding-top: 10px;" align="center">* <strong>Popular List </strong></td>
        </tr>		
<% for k=1 to 3
if not hit.eof Then%>
        <tr> 
          <td align="center" height="8"><img src="title_bar02.htm" width="212" height="5"></td>
        </tr>
        <tr> 
          <td align="center"> 
            <table style="display:{disp1}" border="0" cellpadding="0" cellspacing="0" width="220">
              <tbody><tr> 
                <td align="right" width="78"><table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tbody><tr> 
                      <td align="right" valign="bottom" height="80"><table border="0" cellpadding="0" cellspacing="0" width="72" height="72">
                          <tbody><tr> 
                            <td align="center" background="images/table1_01.gif"><a href="javascript:detail_go(<%=hit("itemkodu")%>);"><img src="item/<%=resim(hit("resim"))%>" border="0" width="64" height="64"></a></td>
                          </tr>
                        </tbody></table></td>
                    </tr>
                    <tr> 
                      <td align="right" valign="bottom" height="32"><a href="javascript:detail_go(<%=hit("itemkodu")%>);" onMouseOver="MM_swapImage('butt_detail1<%=k%>','','images/butt/butt_detail_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt/butt_detail.gif" name="butt_detail1<%=k%>" id="butt_detail1<%=k%>" border="0" width="72" height="30"></a></td>
                    </tr>
                  </tbody></table></td>
                <td align="center" valign="bottom"><table border="0" cellpadding="0" cellspacing="0" width="130" height="108">
                    <tbody><tr> 
                      <td align="right"><table border="0" cellpadding="0" cellspacing="0" width="125">
                          <tbody><tr> 
                            <td width="10">-</td>
                            <td style="font-family: 'verdana',Times,serif; color: rgb(247, 231, 33); font-size: 10px;" height="16"><strong><%=hit("itemismi")%></strong></td>
                          </tr>
                          <tr> 
                            <td>-</td>
                            <td style="font-family: 'verdana',Times,serif; color: rgb(255, 255, 255); font-size: 10px;" height="16">Ucret : <%=hit("ucret")%></td>
                          </tr>
                        </tbody></table></td>
                    </tr>
                    <tr> 
                      <td align="center" height="30"><a href="javascript:buy_go(<%=hit("itemkodu")%>);" onMouseOver="MM_swapImage('butt_buy1<%=k%>','','images/butt/butt_buy_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt/butt_buy.gif" name="butt_buy1<%=k%>" id="butt_buy1<%=k%>" border="0" width="89" height="30"></a></td>
                    </tr>
                    <tr>
                        <td height="2"></td><td>
                    </td></tr>
                    <tr> 
                      <td align="center" height="30"><a href="javascript:wish_go(<%=hit("itemkodu")%>);" onMouseOver="MM_swapImage('butt_cart1<%=k%>','','images/butt/butt_cart_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="images/butt/butt_cart.gif" name="butt_cart1<%=k%>" id="butt_cart1<%=k%>" border="0" width="89" height="30"></a></td>
                    </tr>
                  </tbody></table></td>
              </tr>
            </tbody></table>
          </td>
        </tr>
<% 
hit.movenext
End If
next%>
			
       <tr> 
          <td align="center" height="8">&nbsp;</td>
        </tr>
        <tr> 
          <td align="center">      </td>
        </tr>

                               
            
        

      </tbody></table>
    </td>
  </tr>
  <tr> 
  
    <td style="padding-top: 10px;" align="center">
      <table border="0" cellpadding="0" cellspacing="0" width="252" height="26"><tbody><tr><td style="background-image: url(images/n_table2_02.gif); background-repeat: no-repeat;" height="26">
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tbody><tr>
  <td style="font-size: 11px; color: rgb(255, 238, 47); font-weight: bold;" align="right"><strong><%if not cashp.eof Then
  Response.Write cashp("CashPoint")
  End If%> P</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
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
</body></html><%Else
Response.Write "<br><b>Bu b�l�m Admin tarafindan kapatilmi�tir.</b>"
End If
Menuayar.Close
Set MenuAyar=Nothing%>